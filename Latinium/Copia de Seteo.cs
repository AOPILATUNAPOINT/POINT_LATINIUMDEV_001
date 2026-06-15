using System;
using System.Drawing;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Seteo.
	/// </summary>
	public class Seteo : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnEmpresa;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor16;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.Misc.UltraButton btAutFactura;
		private Infragistics.Win.Misc.UltraButton btSerieFactura;
		private Infragistics.Win.Misc.UltraButton btAutRetencion;
		private Infragistics.Win.Misc.UltraButton btFechaCaduca;
		private Infragistics.Win.Misc.UltraButton btSerieRetencion;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor SriSerieRet;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor SriAutRet;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor SriAutFact;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor SriSerieFact;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDepura;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor13;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor14;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor15;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor5;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor6;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFormMarzo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.Misc.UltraButton btAnexoTrans;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor7;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor9;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor17;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor18;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor19;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutoImpresor;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor20;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteres1;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteres2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor ultraDateTimeEditor1;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor21;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor22;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label lblFiduciario;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFiduciario;
		private System.Windows.Forms.Label lblProyecto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProyecto;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFechaREmi;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFechaRcad;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.Label label35;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCanREmi;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCanRMin;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor23;
		private System.Windows.Forms.Label label37;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor ultraDateTimeEditor2;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor24;
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor25;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpresa;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRUC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAgenteDeRetencion;
		private System.Windows.Forms.Label lblASRetencion;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNResolucion;
		private System.Windows.Forms.Label label43;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEProvincia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEParroquia;
		private System.Windows.Forms.Label label47;
		private System.Windows.Forms.Label label48;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECanton;
		private System.Windows.Forms.Label label49;
		private System.Windows.Forms.Label label50;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalle;
		private System.Windows.Forms.Label label44;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label45;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContClave;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLicenciaContador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContIdentificacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContadorNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label51;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor CedPasRep;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label52;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaResolucion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacionRL;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRepresentanteRL;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label label54;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferenciaRL;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInterseccionRL;
		private System.Windows.Forms.Label label56;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroRL;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProvinciaRL;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbParroquiaRL;
		private System.Windows.Forms.Label label57;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCantonRL;
		private System.Windows.Forms.Label label59;
		private System.Windows.Forms.Label label60;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelularRL;
		private System.Windows.Forms.Label label61;
		private System.Windows.Forms.Label label62;
		private System.Windows.Forms.Label label63;
		private System.Windows.Forms.Label label64;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMailRL;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConvencionalRL;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPrincipalRL;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label65;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource26;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNENacionalidad;
		private C1.Data.C1DataSet cdsSocios;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl8;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private C1.Data.C1DataSet cdsContactos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEmpresasPublicas;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPeriodoMes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl Documentos;
		private System.Windows.Forms.Label label58;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaInscripcion;
		private System.Windows.Forms.Label label66;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicioActividades;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkObligadoLlevarContabilidad;
		private System.Windows.Forms.Label label67;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtActividadEconomica;
		private System.Windows.Forms.Label label68;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbObligaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoEmpresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCumplimientos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSocios;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridContactos;
		private System.Windows.Forms.Label label69;
		private System.Windows.Forms.Label label70;
		private System.Windows.Forms.Label label71;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl11;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridRecordatorios;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Label label72;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRecordatorios;
		private System.Windows.Forms.Button btnUsuarios;
		private System.ComponentModel.IContainer components;

		public Seteo()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Seteo));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAgenteRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNacionalidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNacionalidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Socios", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSocio");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVisa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Socio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNacionalidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomicilioTributario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Participacion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Contactos", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convencional");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Extension");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EMail");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCumplimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresasPublicas", -1, "cmbEmpresasPublicas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdPlanificacion", -1, "cmbPeriodoMes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fechas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFechas", 0);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCumplimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresasPublicas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdPlanificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fechas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdPlanificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdPlanificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObligaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idObligaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presentacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresasPublicas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresasPublicas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVigencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaInscripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaActualizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSubir", 0);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVigencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaConstitucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaInscripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaActualizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRecordatorio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCumplimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioCumplimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSubir", 0);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 1);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRecordatorio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCumplimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioCumplimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab12 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label68 = new System.Windows.Forms.Label();
			this.txtActividadEconomica = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.label67 = new System.Windows.Forms.Label();
			this.cmbTipoEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label66 = new System.Windows.Forms.Label();
			this.dtInicioActividades = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label58 = new System.Windows.Forms.Label();
			this.dtFechaInscripcion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkObligadoLlevarContabilidad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label52 = new System.Windows.Forms.Label();
			this.label51 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label46 = new System.Windows.Forms.Label();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label45 = new System.Windows.Forms.Label();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label44 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbNEProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbNEBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbNEParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label47 = new System.Windows.Forms.Label();
			this.label48 = new System.Windows.Forms.Label();
			this.cmbNECanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label49 = new System.Windows.Forms.Label();
			this.label50 = new System.Windows.Forms.Label();
			this.label43 = new System.Windows.Forms.Label();
			this.dtFechaResolucion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label42 = new System.Windows.Forms.Label();
			this.txtNResolucion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbAgenteDeRetencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblASRetencion = new System.Windows.Forms.Label();
			this.label41 = new System.Windows.Forms.Label();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor23 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label37 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.txtContClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtLicenciaContador = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtContIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtContadorNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRUC = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbNENacionalidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource26 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label53 = new System.Windows.Forms.Label();
			this.label54 = new System.Windows.Forms.Label();
			this.txtReferenciaRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label55 = new System.Windows.Forms.Label();
			this.txtInterseccionRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label56 = new System.Windows.Forms.Label();
			this.txtNumeroRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbProvinciaRL = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbParroquiaRL = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label57 = new System.Windows.Forms.Label();
			this.cmbCantonRL = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label59 = new System.Windows.Forms.Label();
			this.label60 = new System.Windows.Forms.Label();
			this.txtCelularRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label61 = new System.Windows.Forms.Label();
			this.label62 = new System.Windows.Forms.Label();
			this.label63 = new System.Windows.Forms.Label();
			this.label64 = new System.Windows.Forms.Label();
			this.txtEMailRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtConvencionalRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPrincipalRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.CedPasRep = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor12 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtIdentificacionRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtRepresentanteRL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label65 = new System.Windows.Forms.Label();
			this.ultraTabPageControl8 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridSocios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSocios = new C1.Data.C1DataSet();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridContactos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsContactos = new C1.Data.C1DataSet();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridCumplimientos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPeriodoMes = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbObligaciones = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEmpresasPublicas = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.Documentos = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl11 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnUsuarios = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label72 = new System.Windows.Forms.Label();
			this.dtFechaRecordatorios = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridRecordatorios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label71 = new System.Windows.Forms.Label();
			this.label70 = new System.Windows.Forms.Label();
			this.label69 = new System.Windows.Forms.Label();
			this.label38 = new System.Windows.Forms.Label();
			this.ultraDateTimeEditor2 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.ultraTextEditor24 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label40 = new System.Windows.Forms.Label();
			this.ultraTextEditor25 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label36 = new System.Windows.Forms.Label();
			this.label35 = new System.Windows.Forms.Label();
			this.label34 = new System.Windows.Forms.Label();
			this.label33 = new System.Windows.Forms.Label();
			this.cmbFechaRcad = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbFechaREmi = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.spnCanRMin = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnCanREmi = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDateTimeEditor1 = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.ultraTextEditor21 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.ultraTextEditor22 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btSerieRetencion = new Infragistics.Win.Misc.UltraButton();
			this.btFechaCaduca = new Infragistics.Win.Misc.UltraButton();
			this.btAutRetencion = new Infragistics.Win.Misc.UltraButton();
			this.btSerieFactura = new Infragistics.Win.Misc.UltraButton();
			this.btAutFactura = new Infragistics.Win.Misc.UltraButton();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.SriSerieRet = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.SriAutRet = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.ultraTextEditor16 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.SriAutFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.SriSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.ultraTextEditor10 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.spnEmpresa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkDepura = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraCheckEditor10 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label27 = new System.Windows.Forms.Label();
			this.ultraTextEditor20 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.ultraTextEditor15 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label21 = new System.Windows.Forms.Label();
			this.ultraTextEditor14 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.ultraTextEditor13 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lblProyecto = new System.Windows.Forms.Label();
			this.txtProyecto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblFiduciario = new System.Windows.Forms.Label();
			this.txtFiduciario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraCheckEditor9 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCheckEditor8 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraCheckEditor11 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCheckEditor7 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAnexoTrans = new Infragistics.Win.Misc.UltraButton();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFormMarzo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor6 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label29 = new System.Windows.Forms.Label();
			this.txtInteres2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label28 = new System.Windows.Forms.Label();
			this.txtInteres1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkAutoImpresor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label25 = new System.Windows.Forms.Label();
			this.ultraTextEditor19 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.ultraTextEditor18 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.ultraTextEditor17 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtActividadEconomica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicioActividades)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaInscripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaResolucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNResolucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor23)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLicenciaContador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadorNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbNENacionalidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferenciaRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInterseccionRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParroquiaRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelularRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMailRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConvencionalRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrincipalRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.CedPasRep)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacionRL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentanteRL)).BeginInit();
			this.ultraTabPageControl8.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridSocios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSocios)).BeginInit();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridContactos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContactos)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridCumplimientos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodoMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObligaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresasPublicas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.Documentos.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			this.ultraTabPageControl11.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRecordatorios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRecordatorios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraDateTimeEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor24)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor25)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaRcad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaREmi)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCanRMin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCanREmi)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDateTimeEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor22)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.SriSerieRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.SriAutRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.SriAutFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.SriSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEmpresa)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor13)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFiduciario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.label68);
			this.ultraTabPageControl2.Controls.Add(this.txtActividadEconomica);
			this.ultraTabPageControl2.Controls.Add(this.label67);
			this.ultraTabPageControl2.Controls.Add(this.cmbTipoEmpresa);
			this.ultraTabPageControl2.Controls.Add(this.label66);
			this.ultraTabPageControl2.Controls.Add(this.dtInicioActividades);
			this.ultraTabPageControl2.Controls.Add(this.label58);
			this.ultraTabPageControl2.Controls.Add(this.dtFechaInscripcion);
			this.ultraTabPageControl2.Controls.Add(this.chkObligadoLlevarContabilidad);
			this.ultraTabPageControl2.Controls.Add(this.label52);
			this.ultraTabPageControl2.Controls.Add(this.label51);
			this.ultraTabPageControl2.Controls.Add(this.groupBox2);
			this.ultraTabPageControl2.Controls.Add(this.groupBox1);
			this.ultraTabPageControl2.Controls.Add(this.label46);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor4);
			this.ultraTabPageControl2.Controls.Add(this.label45);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl2.Controls.Add(this.label44);
			this.ultraTabPageControl2.Controls.Add(this.txtNumero);
			this.ultraTabPageControl2.Controls.Add(this.cmbNEProvincia);
			this.ultraTabPageControl2.Controls.Add(this.cmbNEBarrio);
			this.ultraTabPageControl2.Controls.Add(this.cmbNEParroquia);
			this.ultraTabPageControl2.Controls.Add(this.label47);
			this.ultraTabPageControl2.Controls.Add(this.label48);
			this.ultraTabPageControl2.Controls.Add(this.cmbNECanton);
			this.ultraTabPageControl2.Controls.Add(this.label49);
			this.ultraTabPageControl2.Controls.Add(this.label50);
			this.ultraTabPageControl2.Controls.Add(this.label43);
			this.ultraTabPageControl2.Controls.Add(this.dtFechaResolucion);
			this.ultraTabPageControl2.Controls.Add(this.label42);
			this.ultraTabPageControl2.Controls.Add(this.txtNResolucion);
			this.ultraTabPageControl2.Controls.Add(this.cmbAgenteDeRetencion);
			this.ultraTabPageControl2.Controls.Add(this.lblASRetencion);
			this.ultraTabPageControl2.Controls.Add(this.label41);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor23);
			this.ultraTabPageControl2.Controls.Add(this.label37);
			this.ultraTabPageControl2.Controls.Add(this.label26);
			this.ultraTabPageControl2.Controls.Add(this.txtContClave);
			this.ultraTabPageControl2.Controls.Add(this.label12);
			this.ultraTabPageControl2.Controls.Add(this.txtLicenciaContador);
			this.ultraTabPageControl2.Controls.Add(this.label9);
			this.ultraTabPageControl2.Controls.Add(this.label7);
			this.ultraTabPageControl2.Controls.Add(this.label6);
			this.ultraTabPageControl2.Controls.Add(this.label4);
			this.ultraTabPageControl2.Controls.Add(this.label3);
			this.ultraTabPageControl2.Controls.Add(this.label2);
			this.ultraTabPageControl2.Controls.Add(this.label1);
			this.ultraTabPageControl2.Controls.Add(this.txtEMail);
			this.ultraTabPageControl2.Controls.Add(this.txtContIdentificacion);
			this.ultraTabPageControl2.Controls.Add(this.txtContadorNombre);
			this.ultraTabPageControl2.Controls.Add(this.txtRUC);
			this.ultraTabPageControl2.Controls.Add(this.txtTelefono);
			this.ultraTabPageControl2.Controls.Add(this.txtCalle);
			this.ultraTabPageControl2.Controls.Add(this.txtEmpresa);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(956, 518);
			// 
			// label68
			// 
			this.label68.AutoSize = true;
			this.label68.BackColor = System.Drawing.Color.Transparent;
			this.label68.Location = new System.Drawing.Point(16, 182);
			this.label68.Name = "label68";
			this.label68.Size = new System.Drawing.Size(106, 17);
			this.label68.TabIndex = 282;
			this.label68.Text = "Actividad Economica";
			// 
			// txtActividadEconomica
			// 
			this.txtActividadEconomica.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtActividadEconomica.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Actividad"));
			this.txtActividadEconomica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtActividadEconomica.Location = new System.Drawing.Point(136, 168);
			this.txtActividadEconomica.Multiline = true;
			this.txtActividadEconomica.Name = "txtActividadEconomica";
			this.txtActividadEconomica.Size = new System.Drawing.Size(800, 44);
			this.txtActividadEconomica.TabIndex = 281;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Location = new System.Drawing.Point(16, 11);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(90, 17);
			this.label67.TabIndex = 280;
			this.label67.Text = "Tipo De Empresa";
			// 
			// cmbTipoEmpresa
			// 
			this.cmbTipoEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoEmpresa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idTipoEmpresa"));
			this.cmbTipoEmpresa.DataSource = this.ultraDataSource4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 296;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoEmpresa.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbTipoEmpresa.DisplayMember = "Descripcion";
			this.cmbTipoEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoEmpresa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoEmpresa.Location = new System.Drawing.Point(136, 8);
			this.cmbTipoEmpresa.Name = "cmbTipoEmpresa";
			this.cmbTipoEmpresa.Size = new System.Drawing.Size(296, 22);
			this.cmbTipoEmpresa.TabIndex = 279;
			this.cmbTipoEmpresa.ValueMember = "idTipoEmpresa";
			this.cmbTipoEmpresa.ValueChanged += new System.EventHandler(this.cmbTipoEmpresa_ValueChanged);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.BackColor = System.Drawing.Color.Transparent;
			this.label66.Location = new System.Drawing.Point(712, 139);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(91, 17);
			this.label66.TabIndex = 278;
			this.label66.Text = "Inicio Actividades";
			// 
			// dtInicioActividades
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicioActividades.Appearance = appearance1;
			this.dtInicioActividades.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtInicioActividades.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.FechaInicio"));
			dateButton1.Caption = "Today";
			this.dtInicioActividades.DateButtons.Add(dateButton1);
			this.dtInicioActividades.Format = "dd/MM/yyyy";
			this.dtInicioActividades.Location = new System.Drawing.Point(824, 137);
			this.dtInicioActividades.Name = "dtInicioActividades";
			this.dtInicioActividades.NonAutoSizeHeight = 23;
			this.dtInicioActividades.Size = new System.Drawing.Size(112, 21);
			this.dtInicioActividades.SpinButtonsVisible = true;
			this.dtInicioActividades.TabIndex = 277;
			this.dtInicioActividades.Value = ((object)(resources.GetObject("dtInicioActividades.Value")));
			// 
			// label58
			// 
			this.label58.AutoSize = true;
			this.label58.BackColor = System.Drawing.Color.Transparent;
			this.label58.Location = new System.Drawing.Point(712, 107);
			this.label58.Name = "label58";
			this.label58.Size = new System.Drawing.Size(91, 17);
			this.label58.TabIndex = 276;
			this.label58.Text = "Fecha Inscripción";
			// 
			// dtFechaInscripcion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaInscripcion.Appearance = appearance2;
			this.dtFechaInscripcion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaInscripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.FechaInscripcion"));
			dateButton2.Caption = "Today";
			this.dtFechaInscripcion.DateButtons.Add(dateButton2);
			this.dtFechaInscripcion.Format = "dd/MM/yyyy";
			this.dtFechaInscripcion.Location = new System.Drawing.Point(824, 105);
			this.dtFechaInscripcion.Name = "dtFechaInscripcion";
			this.dtFechaInscripcion.NonAutoSizeHeight = 23;
			this.dtFechaInscripcion.Size = new System.Drawing.Size(112, 21);
			this.dtFechaInscripcion.SpinButtonsVisible = true;
			this.dtFechaInscripcion.TabIndex = 275;
			this.dtFechaInscripcion.Value = ((object)(resources.GetObject("dtFechaInscripcion.Value")));
			// 
			// chkObligadoLlevarContabilidad
			// 
			appearance3.FontData.Name = "Tahoma";
			this.chkObligadoLlevarContabilidad.Appearance = appearance3;
			this.chkObligadoLlevarContabilidad.BackColor = System.Drawing.Color.Transparent;
			this.chkObligadoLlevarContabilidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkObligadoLlevarContabilidad.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.OblLLevarCont"));
			this.chkObligadoLlevarContabilidad.Location = new System.Drawing.Point(712, 72);
			this.chkObligadoLlevarContabilidad.Name = "chkObligadoLlevarContabilidad";
			this.chkObligadoLlevarContabilidad.Size = new System.Drawing.Size(176, 22);
			this.chkObligadoLlevarContabilidad.TabIndex = 274;
			this.chkObligadoLlevarContabilidad.Text = "Obligado a Llevar Contabilidad";
			// 
			// label52
			// 
			this.label52.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label52.BackColor = System.Drawing.Color.Transparent;
			this.label52.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label52.Location = new System.Drawing.Point(8, 232);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(944, 16);
			this.label52.TabIndex = 273;
			this.label52.Text = "UBICACIÓN DE LA MATRIZ";
			this.label52.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label51
			// 
			this.label51.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label51.BackColor = System.Drawing.Color.Transparent;
			this.label51.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label51.Location = new System.Drawing.Point(8, 424);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(944, 16);
			this.label51.TabIndex = 272;
			this.label51.Text = "DATOS DEL CONTADOR";
			this.label51.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.BackColor = System.Drawing.Color.White;
			this.groupBox2.Location = new System.Drawing.Point(0, 408);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(960, 8);
			this.groupBox2.TabIndex = 271;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.BackColor = System.Drawing.Color.White;
			this.groupBox1.Location = new System.Drawing.Point(0, 216);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(960, 8);
			this.groupBox1.TabIndex = 270;
			this.groupBox1.TabStop = false;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(16, 347);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(57, 17);
			this.label46.TabIndex = 268;
			this.label46.Text = "Referencia";
			// 
			// ultraTextEditor4
			// 
			this.ultraTextEditor4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Referencia"));
			this.ultraTextEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor4.Location = new System.Drawing.Point(80, 344);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(848, 22);
			this.ultraTextEditor4.TabIndex = 269;
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Location = new System.Drawing.Point(544, 315);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(65, 17);
			this.label45.TabIndex = 266;
			this.label45.Text = "Intersección";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Interseccion"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Location = new System.Drawing.Point(628, 312);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(300, 22);
			this.ultraTextEditor2.TabIndex = 267;
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Location = new System.Drawing.Point(400, 315);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(44, 17);
			this.label44.TabIndex = 264;
			this.label44.Text = "Numero";
			// 
			// txtNumero
			// 
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(456, 312);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(70, 22);
			this.txtNumero.TabIndex = 265;
			// 
			// cmbNEProvincia
			// 
			this.cmbNEProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEProvincia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idProvincia"));
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbNEProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNEProvincia.DisplayMember = "Nombre";
			this.cmbNEProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEProvincia.Location = new System.Drawing.Point(16, 280);
			this.cmbNEProvincia.Name = "cmbNEProvincia";
			this.cmbNEProvincia.Size = new System.Drawing.Size(208, 22);
			this.cmbNEProvincia.TabIndex = 263;
			this.cmbNEProvincia.ValueMember = "idProvincia";
			this.cmbNEProvincia.ValueChanged += new System.EventHandler(this.cmbNEProvincia_ValueChanged);
			// 
			// cmbNEBarrio
			// 
			this.cmbNEBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEBarrio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idBarrio"));
			ultraGridBand3.AddButtonCaption = "RolSucursal";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 208;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbNEBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNEBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNEBarrio.DisplayMember = "Nombre";
			this.cmbNEBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEBarrio.Location = new System.Drawing.Point(664, 280);
			this.cmbNEBarrio.Name = "cmbNEBarrio";
			this.cmbNEBarrio.Size = new System.Drawing.Size(208, 22);
			this.cmbNEBarrio.TabIndex = 258;
			this.cmbNEBarrio.ValueMember = "idBarrio";
			// 
			// cmbNEParroquia
			// 
			this.cmbNEParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEParroquia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idParroquia"));
			ultraGridBand4.AddButtonCaption = "RolSucursal";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 208;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbNEParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbNEParroquia.DisplayMember = "Nombre";
			this.cmbNEParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEParroquia.Location = new System.Drawing.Point(448, 280);
			this.cmbNEParroquia.Name = "cmbNEParroquia";
			this.cmbNEParroquia.Size = new System.Drawing.Size(208, 22);
			this.cmbNEParroquia.TabIndex = 257;
			this.cmbNEParroquia.ValueMember = "idParroquia";
			this.cmbNEParroquia.ValueChanged += new System.EventHandler(this.cmbNEParroquia_ValueChanged);
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.White;
			this.label47.Location = new System.Drawing.Point(448, 256);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(52, 17);
			this.label47.TabIndex = 256;
			this.label47.Text = "Parroquia";
			this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.BackColor = System.Drawing.Color.White;
			this.label48.Location = new System.Drawing.Point(664, 256);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(34, 17);
			this.label48.TabIndex = 255;
			this.label48.Text = "Barrio";
			this.label48.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNECanton
			// 
			this.cmbNECanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECanton.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idCanton"));
			ultraGridBand5.AddButtonCaption = "RolSucursal";
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 208;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbNECanton.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbNECanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNECanton.DisplayMember = "Nombre";
			this.cmbNECanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECanton.Location = new System.Drawing.Point(232, 280);
			this.cmbNECanton.Name = "cmbNECanton";
			this.cmbNECanton.Size = new System.Drawing.Size(208, 22);
			this.cmbNECanton.TabIndex = 254;
			this.cmbNECanton.ValueMember = "idCanton";
			this.cmbNECanton.ValueChanged += new System.EventHandler(this.cmbNECanton_ValueChanged);
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.BackColor = System.Drawing.Color.White;
			this.label49.Location = new System.Drawing.Point(16, 256);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(49, 17);
			this.label49.TabIndex = 253;
			this.label49.Text = "Provincia";
			this.label49.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.BackColor = System.Drawing.Color.White;
			this.label50.Location = new System.Drawing.Point(232, 256);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(39, 17);
			this.label50.TabIndex = 252;
			this.label50.Text = "Canton";
			this.label50.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.BackColor = System.Drawing.Color.Transparent;
			this.label43.Location = new System.Drawing.Point(712, 42);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(90, 17);
			this.label43.TabIndex = 251;
			this.label43.Text = "Fecha Resolución";
			// 
			// dtFechaResolucion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaResolucion.Appearance = appearance4;
			this.dtFechaResolucion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaResolucion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.FechaResolucion"));
			dateButton3.Caption = "Today";
			this.dtFechaResolucion.DateButtons.Add(dateButton3);
			this.dtFechaResolucion.Format = "dd/MM/yyyy";
			this.dtFechaResolucion.Location = new System.Drawing.Point(824, 40);
			this.dtFechaResolucion.Name = "dtFechaResolucion";
			this.dtFechaResolucion.NonAutoSizeHeight = 23;
			this.dtFechaResolucion.Size = new System.Drawing.Size(112, 21);
			this.dtFechaResolucion.SpinButtonsVisible = true;
			this.dtFechaResolucion.TabIndex = 250;
			this.dtFechaResolucion.Value = ((object)(resources.GetObject("dtFechaResolucion.Value")));
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(456, 42);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(101, 17);
			this.label42.TabIndex = 248;
			this.label42.Text = "Número Resolución";
			// 
			// txtNResolucion
			// 
			this.txtNResolucion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.NResolucion"));
			this.txtNResolucion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNResolucion.Location = new System.Drawing.Point(576, 39);
			this.txtNResolucion.Name = "txtNResolucion";
			this.txtNResolucion.Size = new System.Drawing.Size(112, 22);
			this.txtNResolucion.TabIndex = 249;
			// 
			// cmbAgenteDeRetencion
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAgenteDeRetencion.Appearance = appearance5;
			this.cmbAgenteDeRetencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbAgenteDeRetencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAgenteDeRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idAgenteRetencion"));
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 350;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbAgenteDeRetencion.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbAgenteDeRetencion.DisplayMember = "Nombre";
			this.cmbAgenteDeRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAgenteDeRetencion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbAgenteDeRetencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAgenteDeRetencion.Location = new System.Drawing.Point(136, 40);
			this.cmbAgenteDeRetencion.Name = "cmbAgenteDeRetencion";
			this.cmbAgenteDeRetencion.Size = new System.Drawing.Size(296, 21);
			this.cmbAgenteDeRetencion.TabIndex = 247;
			this.cmbAgenteDeRetencion.ValueMember = "idAgenteRetencion";
			// 
			// lblASRetencion
			// 
			this.lblASRetencion.AutoSize = true;
			this.lblASRetencion.BackColor = System.Drawing.Color.Transparent;
			this.lblASRetencion.Location = new System.Drawing.Point(16, 42);
			this.lblASRetencion.Name = "lblASRetencion";
			this.lblASRetencion.Size = new System.Drawing.Size(109, 17);
			this.lblASRetencion.TabIndex = 246;
			this.lblASRetencion.Text = "Agente De Retención";
			this.lblASRetencion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Location = new System.Drawing.Point(16, 139);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(96, 17);
			this.label41.TabIndex = 60;
			this.label41.Text = "Nombre Comercial";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.NombreComercial"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(136, 136);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(456, 22);
			this.ultraTextEditor1.TabIndex = 61;
			// 
			// ultraTextEditor23
			// 
			this.ultraTextEditor23.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Celular"));
			this.ultraTextEditor23.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor23.Location = new System.Drawing.Point(296, 376);
			this.ultraTextEditor23.Name = "ultraTextEditor23";
			this.ultraTextEditor23.Size = new System.Drawing.Size(120, 22);
			this.ultraTextEditor23.TabIndex = 58;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(240, 379);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(38, 17);
			this.label37.TabIndex = 57;
			this.label37.Text = "Celular";
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(576, 459);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(31, 17);
			this.label26.TabIndex = 55;
			this.label26.Text = "Clave";
			// 
			// txtContClave
			// 
			this.txtContClave.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.ClaveContador"));
			this.txtContClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContClave.Location = new System.Drawing.Point(624, 456);
			this.txtContClave.Name = "txtContClave";
			this.txtContClave.Size = new System.Drawing.Size(144, 22);
			this.txtContClave.TabIndex = 56;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(336, 459);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(44, 17);
			this.label12.TabIndex = 53;
			this.label12.Text = "Licencia";
			// 
			// txtLicenciaContador
			// 
			this.txtLicenciaContador.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.ContadorLicencia"));
			this.txtLicenciaContador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLicenciaContador.Location = new System.Drawing.Point(400, 456);
			this.txtLicenciaContador.Name = "txtLicenciaContador";
			this.txtLicenciaContador.Size = new System.Drawing.Size(152, 22);
			this.txtLicenciaContador.TabIndex = 54;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(440, 379);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(34, 17);
			this.label9.TabIndex = 47;
			this.label9.Text = "E-Mail";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(16, 459);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(131, 17);
			this.label7.TabIndex = 45;
			this.label7.Text = "Número De Identificación";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 491);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(106, 17);
			this.label6.TabIndex = 41;
			this.label6.Text = "Apellidos y Nombres";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(16, 75);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 17);
			this.label4.TabIndex = 37;
			this.label4.Text = "R.U.C.";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 379);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(70, 17);
			this.label3.TabIndex = 35;
			this.label3.Text = "Convencional";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 315);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(46, 17);
			this.label2.TabIndex = 33;
			this.label2.Text = "Principal";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 107);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(67, 17);
			this.label1.TabIndex = 31;
			this.label1.Text = "Razón Social";
			// 
			// txtEMail
			// 
			this.txtEMail.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Email"));
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Location = new System.Drawing.Point(496, 376);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(176, 22);
			this.txtEMail.TabIndex = 48;
			// 
			// txtContIdentificacion
			// 
			this.txtContIdentificacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CedulaContador"));
			this.txtContIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContIdentificacion.Location = new System.Drawing.Point(160, 456);
			this.txtContIdentificacion.MaxLength = 13;
			this.txtContIdentificacion.Name = "txtContIdentificacion";
			this.txtContIdentificacion.Size = new System.Drawing.Size(152, 22);
			this.txtContIdentificacion.TabIndex = 46;
			// 
			// txtContadorNombre
			// 
			this.txtContadorNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtContadorNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Contador"));
			this.txtContadorNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContadorNombre.Location = new System.Drawing.Point(160, 488);
			this.txtContadorNombre.Name = "txtContadorNombre";
			this.txtContadorNombre.Size = new System.Drawing.Size(392, 22);
			this.txtContadorNombre.TabIndex = 42;
			// 
			// txtRUC
			// 
			this.txtRUC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Ruc"));
			this.txtRUC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRUC.Location = new System.Drawing.Point(136, 72);
			this.txtRUC.MaxLength = 13;
			this.txtRUC.Name = "txtRUC";
			this.txtRUC.Size = new System.Drawing.Size(144, 22);
			this.txtRUC.TabIndex = 38;
			// 
			// txtTelefono
			// 
			this.txtTelefono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Telefono"));
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Location = new System.Drawing.Point(104, 376);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 22);
			this.txtTelefono.TabIndex = 36;
			// 
			// txtCalle
			// 
			this.txtCalle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalle.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Calle"));
			this.txtCalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalle.Location = new System.Drawing.Point(80, 312);
			this.txtCalle.Name = "txtCalle";
			this.txtCalle.Size = new System.Drawing.Size(300, 22);
			this.txtCalle.TabIndex = 34;
			// 
			// txtEmpresa
			// 
			this.txtEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEmpresa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Empresa"));
			this.txtEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpresa.Location = new System.Drawing.Point(136, 104);
			this.txtEmpresa.Name = "txtEmpresa";
			this.txtEmpresa.Size = new System.Drawing.Size(456, 22);
			this.txtEmpresa.TabIndex = 32;
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.cmbNENacionalidad);
			this.ultraTabPageControl7.Controls.Add(this.groupBox3);
			this.ultraTabPageControl7.Controls.Add(this.label53);
			this.ultraTabPageControl7.Controls.Add(this.label54);
			this.ultraTabPageControl7.Controls.Add(this.txtReferenciaRL);
			this.ultraTabPageControl7.Controls.Add(this.label55);
			this.ultraTabPageControl7.Controls.Add(this.txtInterseccionRL);
			this.ultraTabPageControl7.Controls.Add(this.label56);
			this.ultraTabPageControl7.Controls.Add(this.txtNumeroRL);
			this.ultraTabPageControl7.Controls.Add(this.cmbProvinciaRL);
			this.ultraTabPageControl7.Controls.Add(this.cmbParroquiaRL);
			this.ultraTabPageControl7.Controls.Add(this.label57);
			this.ultraTabPageControl7.Controls.Add(this.cmbCantonRL);
			this.ultraTabPageControl7.Controls.Add(this.label59);
			this.ultraTabPageControl7.Controls.Add(this.label60);
			this.ultraTabPageControl7.Controls.Add(this.txtCelularRL);
			this.ultraTabPageControl7.Controls.Add(this.label61);
			this.ultraTabPageControl7.Controls.Add(this.label62);
			this.ultraTabPageControl7.Controls.Add(this.label63);
			this.ultraTabPageControl7.Controls.Add(this.label64);
			this.ultraTabPageControl7.Controls.Add(this.txtEMailRL);
			this.ultraTabPageControl7.Controls.Add(this.txtConvencionalRL);
			this.ultraTabPageControl7.Controls.Add(this.txtPrincipalRL);
			this.ultraTabPageControl7.Controls.Add(this.CedPasRep);
			this.ultraTabPageControl7.Controls.Add(this.ultraTextEditor12);
			this.ultraTabPageControl7.Controls.Add(this.label13);
			this.ultraTabPageControl7.Controls.Add(this.label8);
			this.ultraTabPageControl7.Controls.Add(this.txtIdentificacionRL);
			this.ultraTabPageControl7.Controls.Add(this.label5);
			this.ultraTabPageControl7.Controls.Add(this.txtRepresentanteRL);
			this.ultraTabPageControl7.Controls.Add(this.label65);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(956, 518);
			// 
			// cmbNENacionalidad
			// 
			this.cmbNENacionalidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNENacionalidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNENacionalidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idNacionalidadRL"));
			this.cmbNENacionalidad.DataSource = this.ultraDataSource26;
			ultraGridBand7.AddButtonCaption = "RolSucursal";
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 300;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbNENacionalidad.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbNENacionalidad.DisplayMember = "Nombre";
			this.cmbNENacionalidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNENacionalidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNENacionalidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNENacionalidad.Location = new System.Drawing.Point(664, 40);
			this.cmbNENacionalidad.Name = "cmbNENacionalidad";
			this.cmbNENacionalidad.Size = new System.Drawing.Size(216, 22);
			this.cmbNENacionalidad.TabIndex = 354;
			this.cmbNENacionalidad.ValueMember = "idNacionalidad";
			// 
			// ultraDataSource26
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource26.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn3,
																																	ultraDataColumn4});
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.BackColor = System.Drawing.Color.Transparent;
			this.groupBox3.Location = new System.Drawing.Point(0, 72);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(960, 8);
			this.groupBox3.TabIndex = 297;
			this.groupBox3.TabStop = false;
			// 
			// label53
			// 
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label53.Location = new System.Drawing.Point(8, 88);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(944, 16);
			this.label53.TabIndex = 296;
			this.label53.Text = "UBICACIÓN DEL DOMICILIO";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Location = new System.Drawing.Point(16, 203);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(57, 17);
			this.label54.TabIndex = 294;
			this.label54.Text = "Referencia";
			// 
			// txtReferenciaRL
			// 
			this.txtReferenciaRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtReferenciaRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.ReferenciaRL"));
			this.txtReferenciaRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferenciaRL.Location = new System.Drawing.Point(79, 200);
			this.txtReferenciaRL.Name = "txtReferenciaRL";
			this.txtReferenciaRL.Size = new System.Drawing.Size(853, 22);
			this.txtReferenciaRL.TabIndex = 295;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Location = new System.Drawing.Point(544, 171);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(65, 17);
			this.label55.TabIndex = 292;
			this.label55.Text = "Intersección";
			// 
			// txtInterseccionRL
			// 
			this.txtInterseccionRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInterseccionRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.InterseccionRL"));
			this.txtInterseccionRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInterseccionRL.Location = new System.Drawing.Point(632, 168);
			this.txtInterseccionRL.Name = "txtInterseccionRL";
			this.txtInterseccionRL.Size = new System.Drawing.Size(300, 22);
			this.txtInterseccionRL.TabIndex = 293;
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.BackColor = System.Drawing.Color.Transparent;
			this.label56.Location = new System.Drawing.Point(400, 171);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(44, 17);
			this.label56.TabIndex = 290;
			this.label56.Text = "Numero";
			// 
			// txtNumeroRL
			// 
			this.txtNumeroRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumeroRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.NumeroRL"));
			this.txtNumeroRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroRL.Location = new System.Drawing.Point(456, 168);
			this.txtNumeroRL.Name = "txtNumeroRL";
			this.txtNumeroRL.Size = new System.Drawing.Size(70, 22);
			this.txtNumeroRL.TabIndex = 291;
			// 
			// cmbProvinciaRL
			// 
			this.cmbProvinciaRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProvinciaRL.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvinciaRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idProvinciaRL"));
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 208;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbProvinciaRL.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbProvinciaRL.DisplayMember = "Nombre";
			this.cmbProvinciaRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProvinciaRL.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProvinciaRL.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvinciaRL.Location = new System.Drawing.Point(16, 136);
			this.cmbProvinciaRL.Name = "cmbProvinciaRL";
			this.cmbProvinciaRL.Size = new System.Drawing.Size(208, 22);
			this.cmbProvinciaRL.TabIndex = 289;
			this.cmbProvinciaRL.ValueMember = "idProvincia";
			this.cmbProvinciaRL.ValueChanged += new System.EventHandler(this.cmbProvinciaRL_ValueChanged);
			// 
			// cmbParroquiaRL
			// 
			this.cmbParroquiaRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbParroquiaRL.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParroquiaRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idParroquiaRL"));
			ultraGridBand9.AddButtonCaption = "RolSucursal";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 208;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbParroquiaRL.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbParroquiaRL.DisplayMember = "Nombre";
			this.cmbParroquiaRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbParroquiaRL.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbParroquiaRL.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParroquiaRL.Location = new System.Drawing.Point(448, 136);
			this.cmbParroquiaRL.Name = "cmbParroquiaRL";
			this.cmbParroquiaRL.Size = new System.Drawing.Size(208, 22);
			this.cmbParroquiaRL.TabIndex = 287;
			this.cmbParroquiaRL.ValueMember = "idParroquia";
			// 
			// label57
			// 
			this.label57.AutoSize = true;
			this.label57.BackColor = System.Drawing.Color.White;
			this.label57.Location = new System.Drawing.Point(448, 112);
			this.label57.Name = "label57";
			this.label57.Size = new System.Drawing.Size(52, 17);
			this.label57.TabIndex = 286;
			this.label57.Text = "Parroquia";
			this.label57.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCantonRL
			// 
			this.cmbCantonRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCantonRL.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCantonRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.idCantonRL"));
			ultraGridBand10.AddButtonCaption = "RolSucursal";
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 208;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn19,
																											ultraGridColumn20});
			this.cmbCantonRL.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbCantonRL.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCantonRL.DisplayMember = "Nombre";
			this.cmbCantonRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCantonRL.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCantonRL.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCantonRL.Location = new System.Drawing.Point(232, 136);
			this.cmbCantonRL.Name = "cmbCantonRL";
			this.cmbCantonRL.Size = new System.Drawing.Size(208, 22);
			this.cmbCantonRL.TabIndex = 284;
			this.cmbCantonRL.ValueMember = "idCanton";
			this.cmbCantonRL.ValueChanged += new System.EventHandler(this.cmbCantonRL_ValueChanged);
			// 
			// label59
			// 
			this.label59.AutoSize = true;
			this.label59.BackColor = System.Drawing.Color.White;
			this.label59.Location = new System.Drawing.Point(16, 112);
			this.label59.Name = "label59";
			this.label59.Size = new System.Drawing.Size(49, 17);
			this.label59.TabIndex = 283;
			this.label59.Text = "Provincia";
			this.label59.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label60
			// 
			this.label60.AutoSize = true;
			this.label60.BackColor = System.Drawing.Color.White;
			this.label60.Location = new System.Drawing.Point(232, 112);
			this.label60.Name = "label60";
			this.label60.Size = new System.Drawing.Size(39, 17);
			this.label60.TabIndex = 282;
			this.label60.Text = "Canton";
			this.label60.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCelularRL
			// 
			this.txtCelularRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CelularRL"));
			this.txtCelularRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelularRL.Location = new System.Drawing.Point(296, 232);
			this.txtCelularRL.Name = "txtCelularRL";
			this.txtCelularRL.Size = new System.Drawing.Size(120, 22);
			this.txtCelularRL.TabIndex = 281;
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(240, 235);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(38, 17);
			this.label61.TabIndex = 280;
			this.label61.Text = "Celular";
			// 
			// label62
			// 
			this.label62.AutoSize = true;
			this.label62.BackColor = System.Drawing.Color.Transparent;
			this.label62.Location = new System.Drawing.Point(440, 235);
			this.label62.Name = "label62";
			this.label62.Size = new System.Drawing.Size(34, 17);
			this.label62.TabIndex = 278;
			this.label62.Text = "E-Mail";
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.BackColor = System.Drawing.Color.Transparent;
			this.label63.Location = new System.Drawing.Point(16, 235);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(70, 17);
			this.label63.TabIndex = 276;
			this.label63.Text = "Convencional";
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.BackColor = System.Drawing.Color.Transparent;
			this.label64.Location = new System.Drawing.Point(16, 171);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(46, 17);
			this.label64.TabIndex = 274;
			this.label64.Text = "Principal";
			// 
			// txtEMailRL
			// 
			this.txtEMailRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.EmailRL"));
			this.txtEMailRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMailRL.Location = new System.Drawing.Point(496, 232);
			this.txtEMailRL.Name = "txtEMailRL";
			this.txtEMailRL.Size = new System.Drawing.Size(176, 22);
			this.txtEMailRL.TabIndex = 279;
			// 
			// txtConvencionalRL
			// 
			this.txtConvencionalRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.ConvencionalRL"));
			this.txtConvencionalRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConvencionalRL.Location = new System.Drawing.Point(104, 232);
			this.txtConvencionalRL.Name = "txtConvencionalRL";
			this.txtConvencionalRL.Size = new System.Drawing.Size(112, 22);
			this.txtConvencionalRL.TabIndex = 277;
			// 
			// txtPrincipalRL
			// 
			this.txtPrincipalRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPrincipalRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CalleRL"));
			this.txtPrincipalRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrincipalRL.Location = new System.Drawing.Point(80, 168);
			this.txtPrincipalRL.Name = "txtPrincipalRL";
			this.txtPrincipalRL.Size = new System.Drawing.Size(300, 22);
			this.txtPrincipalRL.TabIndex = 275;
			// 
			// CedPasRep
			// 
			this.CedPasRep.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CedPasRep"));
			this.CedPasRep.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.CedPasRep.Location = new System.Drawing.Point(640, 304);
			this.CedPasRep.Name = "CedPasRep";
			this.CedPasRep.Size = new System.Drawing.Size(24, 22);
			this.CedPasRep.TabIndex = 62;
			this.CedPasRep.Text = "C/P";
			this.toolTip1.SetToolTip(this.CedPasRep, "Identificacion de Representante C Cedula P pasaporte");
			this.CedPasRep.Visible = false;
			// 
			// ultraTextEditor12
			// 
			this.ultraTextEditor12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Fax"));
			this.ultraTextEditor12.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor12.Location = new System.Drawing.Point(712, 304);
			this.ultraTextEditor12.Name = "ultraTextEditor12";
			this.ultraTextEditor12.Size = new System.Drawing.Size(120, 22);
			this.ultraTextEditor12.TabIndex = 61;
			this.ultraTextEditor12.Visible = false;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(672, 304);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(26, 17);
			this.label13.TabIndex = 60;
			this.label13.Text = "Fax:";
			this.label13.Visible = false;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(8, 11);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(71, 17);
			this.label8.TabIndex = 45;
			this.label8.Text = "Identificación";
			// 
			// txtIdentificacionRL
			// 
			this.txtIdentificacionRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CedulaRep"));
			this.txtIdentificacionRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdentificacionRL.Location = new System.Drawing.Point(96, 8);
			this.txtIdentificacionRL.Name = "txtIdentificacionRL";
			this.txtIdentificacionRL.Size = new System.Drawing.Size(128, 22);
			this.txtIdentificacionRL.TabIndex = 46;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(77, 17);
			this.label5.TabIndex = 41;
			this.label5.Text = "Representante";
			// 
			// txtRepresentanteRL
			// 
			this.txtRepresentanteRL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRepresentanteRL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Representante"));
			this.txtRepresentanteRL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepresentanteRL.Location = new System.Drawing.Point(96, 40);
			this.txtRepresentanteRL.Name = "txtRepresentanteRL";
			this.txtRepresentanteRL.Size = new System.Drawing.Size(456, 22);
			this.txtRepresentanteRL.TabIndex = 42;
			// 
			// label65
			// 
			this.label65.AutoSize = true;
			this.label65.BackColor = System.Drawing.Color.Transparent;
			this.label65.Location = new System.Drawing.Point(576, 43);
			this.label65.Name = "label65";
			this.label65.Size = new System.Drawing.Size(68, 17);
			this.label65.TabIndex = 353;
			this.label65.Text = "Nacionalidad";
			this.label65.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraTabPageControl8
			// 
			this.ultraTabPageControl8.Controls.Add(this.uGridSocios);
			this.ultraTabPageControl8.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl8.Name = "ultraTabPageControl8";
			this.ultraTabPageControl8.Size = new System.Drawing.Size(956, 518);
			// 
			// uGridSocios
			// 
			this.uGridSocios.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSocios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSocios.DataSource = this.cdsSocios;
			appearance6.BackColor = System.Drawing.Color.White;
			this.uGridSocios.DisplayLayout.Appearance = appearance6;
			this.uGridSocios.DisplayLayout.AutoFitColumns = true;
			ultraGridBand11.AddButtonCaption = "ClienteTipoRuc";
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellAppearance = appearance7;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 54;
			ultraGridColumn22.Header.Caption = "Tipo Identificación";
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 61;
			ultraGridColumn23.Header.Caption = "Identificación";
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Width = 110;
			ultraGridColumn24.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn24.Header.Caption = "Tipo Visa";
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Width = 91;
			ultraGridColumn25.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Width = 250;
			ultraGridColumn26.Header.Caption = "Nacionalidad";
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Width = 160;
			ultraGridColumn27.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn27.Header.Caption = "Domicilio Tributario";
			ultraGridColumn27.Header.VisiblePosition = 6;
			ultraGridColumn27.Width = 184;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance8;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "% Participación";
			ultraGridColumn28.Header.VisiblePosition = 7;
			ultraGridColumn28.Width = 120;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn29.CellAppearance = appearance9;
			ultraGridColumn29.Header.VisiblePosition = 8;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 82;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn30.CellAppearance = appearance10;
			ultraGridColumn30.Header.VisiblePosition = 9;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 91;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn31.CellAppearance = appearance11;
			ultraGridColumn31.Header.VisiblePosition = 10;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 97;
			ultraGridBand11.Columns.AddRange(new object[] {
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
																											ultraGridColumn31});
			this.uGridSocios.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.uGridSocios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSocios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSocios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridSocios.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.75F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSocios.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridSocios.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSocios.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSocios.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridSocios.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSocios.Location = new System.Drawing.Point(8, 8);
			this.uGridSocios.Name = "uGridSocios";
			this.uGridSocios.Size = new System.Drawing.Size(936, 504);
			this.uGridSocios.TabIndex = 2;
			this.uGridSocios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSocios_KeyDown);
			// 
			// cdsSocios
			// 
			this.cdsSocios.BindingContextCtrl = this;
			this.cdsSocios.DataLibrary = "LibContabilidad";
			this.cdsSocios.DataLibraryUrl = "";
			this.cdsSocios.DataSetDef = "dsSocios";
			this.cdsSocios.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSocios.Name = "cdsSocios";
			this.cdsSocios.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSocios.SchemaDef = null;
			this.cdsSocios.BeforeFill += new C1.Data.FillEventHandler(this.cdsSocios_BeforeFill);
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.uGridContactos);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(956, 518);
			// 
			// uGridContactos
			// 
			this.uGridContactos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridContactos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridContactos.DataSource = this.cdsContactos;
			appearance16.BackColor = System.Drawing.Color.White;
			this.uGridContactos.DisplayLayout.Appearance = appearance16;
			ultraGridBand12.AddButtonCaption = "ClienteTipoRuc";
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 250;
			ultraGridColumn34.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn34.Width = 250;
			ultraGridColumn35.Header.Caption = "Telefono";
			ultraGridColumn35.Header.VisiblePosition = 3;
			ultraGridColumn35.Width = 80;
			ultraGridColumn36.Header.Caption = "Ext";
			ultraGridColumn36.Header.VisiblePosition = 4;
			ultraGridColumn36.Width = 50;
			ultraGridColumn37.Header.VisiblePosition = 5;
			ultraGridColumn37.Width = 80;
			ultraGridColumn38.Header.Caption = "E-Mail";
			ultraGridColumn38.Header.VisiblePosition = 6;
			ultraGridColumn38.Width = 150;
			ultraGridColumn39.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn39.Header.VisiblePosition = 7;
			ultraGridColumn39.Width = 300;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn40.CellAppearance = appearance17;
			ultraGridColumn40.Header.VisiblePosition = 8;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 82;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn41.CellAppearance = appearance18;
			ultraGridColumn41.Header.VisiblePosition = 9;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 91;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn42.CellAppearance = appearance19;
			ultraGridColumn42.Header.VisiblePosition = 10;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 97;
			ultraGridBand12.Columns.AddRange(new object[] {
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
																											ultraGridColumn42});
			this.uGridContactos.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.uGridContactos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridContactos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridContactos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance20.BackColor = System.Drawing.Color.Transparent;
			this.uGridContactos.DisplayLayout.Override.CardAreaAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance21.FontData.BoldAsString = "True";
			appearance21.FontData.Name = "Arial";
			appearance21.FontData.SizeInPoints = 9F;
			appearance21.ForeColor = System.Drawing.Color.White;
			appearance21.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridContactos.DisplayLayout.Override.HeaderAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContactos.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContactos.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.uGridContactos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridContactos.Location = new System.Drawing.Point(10, 11);
			this.uGridContactos.Name = "uGridContactos";
			this.uGridContactos.Size = new System.Drawing.Size(936, 501);
			this.uGridContactos.TabIndex = 3;
			this.uGridContactos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridContactos_KeyDown);
			// 
			// cdsContactos
			// 
			this.cdsContactos.BindingContextCtrl = this;
			this.cdsContactos.DataLibrary = "LibContabilidad";
			this.cdsContactos.DataLibraryUrl = "";
			this.cdsContactos.DataSetDef = "dsContactos";
			this.cdsContactos.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsContactos.Name = "cdsContactos";
			this.cdsContactos.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsContactos.SchemaDef = null;
			this.cdsContactos.BeforeFill += new C1.Data.FillEventHandler(this.cdsContactos_BeforeFill);
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.uGridCumplimientos);
			this.ultraTabPageControl10.Controls.Add(this.cmbPeriodoMes);
			this.ultraTabPageControl10.Controls.Add(this.cmbObligaciones);
			this.ultraTabPageControl10.Controls.Add(this.cmbEmpresasPublicas);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(956, 518);
			// 
			// uGridCumplimientos
			// 
			this.uGridCumplimientos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCumplimientos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCumplimientos.DataSource = this.ultraDataSource5;
			appearance24.BackColor = System.Drawing.Color.White;
			this.uGridCumplimientos.DisplayLayout.Appearance = appearance24;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.Caption = "...";
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 20;
			ultraGridColumn45.Header.Caption = "Empresas";
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn45.Width = 150;
			ultraGridColumn46.Header.Caption = "Periodos";
			ultraGridColumn46.Header.VisiblePosition = 4;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn46.Width = 120;
			ultraGridColumn47.Header.VisiblePosition = 6;
			ultraGridColumn47.Width = 200;
			ultraGridColumn48.Header.VisiblePosition = 3;
			ultraGridColumn48.Width = 300;
			ultraGridColumn49.Header.VisiblePosition = 7;
			ultraGridColumn49.Width = 500;
			ultraGridColumn50.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance25.Image = ((object)(resources.GetObject("appearance25.Image")));
			appearance25.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance25.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn50.CellButtonAppearance = appearance25;
			ultraGridColumn50.Header.Caption = "...";
			ultraGridColumn50.Header.VisiblePosition = 5;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn50.Width = 30;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn43,
																											ultraGridColumn44,
																											ultraGridColumn45,
																											ultraGridColumn46,
																											ultraGridColumn47,
																											ultraGridColumn48,
																											ultraGridColumn49,
																											ultraGridColumn50});
			this.uGridCumplimientos.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.uGridCumplimientos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridCumplimientos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridCumplimientos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.uGridCumplimientos.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 9F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCumplimientos.DisplayLayout.Override.HeaderAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCumplimientos.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCumplimientos.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridCumplimientos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCumplimientos.Location = new System.Drawing.Point(10, 9);
			this.uGridCumplimientos.Name = "uGridCumplimientos";
			this.uGridCumplimientos.Size = new System.Drawing.Size(936, 501);
			this.uGridCumplimientos.TabIndex = 347;
			this.uGridCumplimientos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCumplimientos_KeyDown);
			this.uGridCumplimientos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCumplimientos_ClickCellButton);
			this.uGridCumplimientos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCumplimientos_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn6.DefaultValue = false;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// cmbPeriodoMes
			// 
			this.cmbPeriodoMes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPeriodoMes.DataSource = this.ultraDataSource2;
			this.cmbPeriodoMes.DisplayLayout.AutoFitColumns = true;
			ultraGridBand14.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 67;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 278;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn51,
																											ultraGridColumn52});
			this.cmbPeriodoMes.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbPeriodoMes.DisplayMember = "Descripcion";
			this.cmbPeriodoMes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPeriodoMes.Location = new System.Drawing.Point(640, 208);
			this.cmbPeriodoMes.Name = "cmbPeriodoMes";
			this.cmbPeriodoMes.Size = new System.Drawing.Size(280, 112);
			this.cmbPeriodoMes.TabIndex = 342;
			this.cmbPeriodoMes.ValueMember = "IdPlanificacion";
			this.cmbPeriodoMes.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// cmbObligaciones
			// 
			this.cmbObligaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbObligaciones.DataSource = this.ultraDataSource3;
			ultraGridBand15.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn53.Header.VisiblePosition = 0;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 37;
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Width = 300;
			ultraGridColumn55.Header.VisiblePosition = 2;
			ultraGridColumn55.Width = 300;
			ultraGridColumn56.Header.VisiblePosition = 3;
			ultraGridColumn56.Width = 300;
			ultraGridColumn57.Header.VisiblePosition = 4;
			ultraGridColumn57.Width = 300;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn53,
																											ultraGridColumn54,
																											ultraGridColumn55,
																											ultraGridColumn56,
																											ultraGridColumn57});
			this.cmbObligaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbObligaciones.DisplayMember = "Descripcion";
			this.cmbObligaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbObligaciones.Location = new System.Drawing.Point(64, 296);
			this.cmbObligaciones.Name = "cmbObligaciones";
			this.cmbObligaciones.Size = new System.Drawing.Size(280, 112);
			this.cmbObligaciones.TabIndex = 343;
			this.cmbObligaciones.ValueMember = "idObligaciones";
			this.cmbObligaciones.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// cmbEmpresasPublicas
			// 
			this.cmbEmpresasPublicas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresasPublicas.DataSource = this.ultraDataSource1;
			this.cmbEmpresasPublicas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand16.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn58.Header.VisiblePosition = 0;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 152;
			ultraGridColumn59.Header.VisiblePosition = 1;
			ultraGridColumn59.Width = 278;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn58,
																											ultraGridColumn59});
			this.cmbEmpresasPublicas.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbEmpresasPublicas.DisplayMember = "Descripcion";
			this.cmbEmpresasPublicas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresasPublicas.Location = new System.Drawing.Point(338, 167);
			this.cmbEmpresasPublicas.Name = "cmbEmpresasPublicas";
			this.cmbEmpresasPublicas.Size = new System.Drawing.Size(280, 112);
			this.cmbEmpresasPublicas.TabIndex = 341;
			this.cmbEmpresasPublicas.ValueMember = "idEmpresasPublicas";
			this.cmbEmpresasPublicas.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// Documentos
			// 
			this.Documentos.Controls.Add(this.uGridDocumentos);
			this.Documentos.Location = new System.Drawing.Point(-10000, -10000);
			this.Documentos.Name = "Documentos";
			this.Documentos.Size = new System.Drawing.Size(956, 518);
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource6;
			appearance30.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance30;
			ultraGridBand17.AddButtonCaption = "Cumplimientos";
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.VisiblePosition = 0;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.Caption = "...";
			ultraGridColumn61.Header.VisiblePosition = 1;
			ultraGridColumn61.Width = 18;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn62.Width = 300;
			ultraGridColumn63.Header.VisiblePosition = 3;
			ultraGridColumn63.Width = 300;
			ultraGridColumn64.Header.Caption = "Vigencia";
			ultraGridColumn64.Header.VisiblePosition = 5;
			ultraGridColumn64.Width = 85;
			ultraGridColumn65.Header.Caption = "Constitucion";
			ultraGridColumn65.Header.VisiblePosition = 6;
			ultraGridColumn65.Width = 85;
			ultraGridColumn66.Header.Caption = "Inscripcion";
			ultraGridColumn66.Header.VisiblePosition = 7;
			ultraGridColumn66.Width = 85;
			ultraGridColumn67.Header.Caption = "Actualizacion";
			ultraGridColumn67.Header.VisiblePosition = 8;
			ultraGridColumn67.Width = 85;
			ultraGridColumn68.Header.VisiblePosition = 4;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 400;
			ultraGridColumn69.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance31.Image = ((object)(resources.GetObject("appearance31.Image")));
			appearance31.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance31.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn69.CellButtonAppearance = appearance31;
			ultraGridColumn69.Header.Caption = "Subir";
			ultraGridColumn69.Header.VisiblePosition = 9;
			ultraGridColumn69.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn69.Width = 40;
			ultraGridColumn70.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance32.Image = ((object)(resources.GetObject("appearance32.Image")));
			appearance32.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance32.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn70.CellButtonAppearance = appearance32;
			ultraGridColumn70.Header.Caption = "Ver";
			ultraGridColumn70.Header.VisiblePosition = 10;
			ultraGridColumn70.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn70.Width = 40;
			ultraGridBand17.Columns.AddRange(new object[] {
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
																											ultraGridColumn70});
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.FontData.BoldAsString = "True";
			appearance34.FontData.Name = "Arial";
			appearance34.FontData.SizeInPoints = 8.5F;
			appearance34.ForeColor = System.Drawing.Color.White;
			appearance34.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(6, 9);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(944, 501);
			this.uGridDocumentos.TabIndex = 5;
			this.uGridDocumentos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDocumentos_KeyDown);
			this.uGridDocumentos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_ClickCellButton);
			this.uGridDocumentos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDocumentos_InitializeLayout);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(bool);
			ultraDataColumn25.DataType = typeof(System.DateTime);
			ultraDataColumn26.DataType = typeof(System.DateTime);
			ultraDataColumn27.DataType = typeof(System.DateTime);
			ultraDataColumn28.DataType = typeof(System.DateTime);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// ultraTabPageControl11
			// 
			this.ultraTabPageControl11.Controls.Add(this.btnUsuarios);
			this.ultraTabPageControl11.Controls.Add(this.btnVer);
			this.ultraTabPageControl11.Controls.Add(this.optFiltro);
			this.ultraTabPageControl11.Controls.Add(this.label72);
			this.ultraTabPageControl11.Controls.Add(this.dtFechaRecordatorios);
			this.ultraTabPageControl11.Controls.Add(this.uGridRecordatorios);
			this.ultraTabPageControl11.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl11.Name = "ultraTabPageControl11";
			this.ultraTabPageControl11.Size = new System.Drawing.Size(956, 518);
			// 
			// btnUsuarios
			// 
			this.btnUsuarios.Image = ((System.Drawing.Image)(resources.GetObject("btnUsuarios.Image")));
			this.btnUsuarios.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnUsuarios.Location = new System.Drawing.Point(872, 6);
			this.btnUsuarios.Name = "btnUsuarios";
			this.btnUsuarios.Size = new System.Drawing.Size(75, 24);
			this.btnUsuarios.TabIndex = 360;
			this.btnUsuarios.Text = "Usuarios";
			this.btnUsuarios.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnUsuarios.Click += new System.EventHandler(this.btnUsuarios_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(376, 6);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 359;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// optFiltro
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance37;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 1;
			this.optFiltro.ItemAppearance = appearance38;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = true;
			valueListItem1.DisplayText = "Cumplido";
			valueListItem2.DataValue = false;
			valueListItem2.DisplayText = "Pendiente";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(192, 6);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(168, 24);
			this.optFiltro.TabIndex = 254;
			this.optFiltro.Text = "Pendiente";
			// 
			// label72
			// 
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Location = new System.Drawing.Point(8, 10);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(34, 17);
			this.label72.TabIndex = 253;
			this.label72.Text = "Fecha";
			// 
			// dtFechaRecordatorios
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRecordatorios.Appearance = appearance39;
			this.dtFechaRecordatorios.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFechaRecordatorios.DateButtons.Add(dateButton4);
			this.dtFechaRecordatorios.Format = "dd/MM/yyyy";
			this.dtFechaRecordatorios.Location = new System.Drawing.Point(64, 8);
			this.dtFechaRecordatorios.Name = "dtFechaRecordatorios";
			this.dtFechaRecordatorios.NonAutoSizeHeight = 23;
			this.dtFechaRecordatorios.Size = new System.Drawing.Size(112, 21);
			this.dtFechaRecordatorios.SpinButtonsVisible = true;
			this.dtFechaRecordatorios.TabIndex = 252;
			this.dtFechaRecordatorios.Value = ((object)(resources.GetObject("dtFechaRecordatorios.Value")));
			// 
			// uGridRecordatorios
			// 
			this.uGridRecordatorios.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRecordatorios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRecordatorios.DataSource = this.ultraDataSource7;
			appearance40.BackColor = System.Drawing.Color.White;
			this.uGridRecordatorios.DisplayLayout.Appearance = appearance40;
			ultraGridBand18.AddButtonCaption = "Cumplimientos";
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn71.Header.VisiblePosition = 0;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn72.Header.VisiblePosition = 1;
			ultraGridColumn72.Width = 70;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn73.Header.VisiblePosition = 2;
			ultraGridColumn73.Width = 300;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn74.Header.VisiblePosition = 3;
			ultraGridColumn74.Width = 300;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn75.Header.VisiblePosition = 4;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn76.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn76.Header.Caption = "Cumplimiento";
			ultraGridColumn76.Header.VisiblePosition = 5;
			ultraGridColumn76.Width = 120;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn77.Header.Caption = "Usuario";
			ultraGridColumn77.Header.VisiblePosition = 6;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn78.Header.VisiblePosition = 7;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance41.Image = ((object)(resources.GetObject("appearance41.Image")));
			appearance41.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance41.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn79.CellButtonAppearance = appearance41;
			ultraGridColumn79.Header.Caption = "Subir";
			ultraGridColumn79.Header.VisiblePosition = 8;
			ultraGridColumn79.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn79.Width = 40;
			ultraGridColumn80.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance42.Image = ((object)(resources.GetObject("appearance42.Image")));
			appearance42.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance42.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn80.CellButtonAppearance = appearance42;
			ultraGridColumn80.Header.Caption = "Ver";
			ultraGridColumn80.Header.VisiblePosition = 9;
			ultraGridColumn80.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn80.Width = 40;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn71,
																											ultraGridColumn72,
																											ultraGridColumn73,
																											ultraGridColumn74,
																											ultraGridColumn75,
																											ultraGridColumn76,
																											ultraGridColumn77,
																											ultraGridColumn78,
																											ultraGridColumn79,
																											ultraGridColumn80});
			this.uGridRecordatorios.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.uGridRecordatorios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridRecordatorios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridRecordatorios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance43.BackColor = System.Drawing.Color.Transparent;
			this.uGridRecordatorios.DisplayLayout.Override.CardAreaAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance44.FontData.BoldAsString = "True";
			appearance44.FontData.Name = "Arial";
			appearance44.FontData.SizeInPoints = 8.5F;
			appearance44.ForeColor = System.Drawing.Color.White;
			appearance44.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRecordatorios.DisplayLayout.Override.HeaderAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecordatorios.DisplayLayout.Override.RowSelectorAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRecordatorios.DisplayLayout.Override.SelectedRowAppearance = appearance46;
			this.uGridRecordatorios.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRecordatorios.Location = new System.Drawing.Point(6, 40);
			this.uGridRecordatorios.Name = "uGridRecordatorios";
			this.uGridRecordatorios.Size = new System.Drawing.Size(944, 464);
			this.uGridRecordatorios.TabIndex = 6;
			this.uGridRecordatorios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridRecordatorios_KeyDown);
			this.uGridRecordatorios.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridRecordatorios_ClickCellButton);
			this.uGridRecordatorios.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridRecordatorios_InitializeLayout);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(System.DateTime);
			ultraDataColumn34.DataType = typeof(bool);
			ultraDataColumn35.DataType = typeof(System.DateTime);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.label71);
			this.ultraTabPageControl1.Controls.Add(this.label70);
			this.ultraTabPageControl1.Controls.Add(this.label69);
			this.ultraTabPageControl1.Controls.Add(this.label38);
			this.ultraTabPageControl1.Controls.Add(this.ultraDateTimeEditor2);
			this.ultraTabPageControl1.Controls.Add(this.label39);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor24);
			this.ultraTabPageControl1.Controls.Add(this.label40);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor25);
			this.ultraTabPageControl1.Controls.Add(this.label36);
			this.ultraTabPageControl1.Controls.Add(this.label35);
			this.ultraTabPageControl1.Controls.Add(this.label34);
			this.ultraTabPageControl1.Controls.Add(this.label33);
			this.ultraTabPageControl1.Controls.Add(this.cmbFechaRcad);
			this.ultraTabPageControl1.Controls.Add(this.cmbFechaREmi);
			this.ultraTabPageControl1.Controls.Add(this.spnCanRMin);
			this.ultraTabPageControl1.Controls.Add(this.spnCanREmi);
			this.ultraTabPageControl1.Controls.Add(this.ultraDateTimeEditor1);
			this.ultraTabPageControl1.Controls.Add(this.label30);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor21);
			this.ultraTabPageControl1.Controls.Add(this.label31);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor22);
			this.ultraTabPageControl1.Controls.Add(this.label32);
			this.ultraTabPageControl1.Controls.Add(this.ultraCheckEditor3);
			this.ultraTabPageControl1.Controls.Add(this.btSerieRetencion);
			this.ultraTabPageControl1.Controls.Add(this.btFechaCaduca);
			this.ultraTabPageControl1.Controls.Add(this.btAutRetencion);
			this.ultraTabPageControl1.Controls.Add(this.btSerieFactura);
			this.ultraTabPageControl1.Controls.Add(this.btAutFactura);
			this.ultraTabPageControl1.Controls.Add(this.cmbFecha);
			this.ultraTabPageControl1.Controls.Add(this.label18);
			this.ultraTabPageControl1.Controls.Add(this.SriSerieRet);
			this.ultraTabPageControl1.Controls.Add(this.label19);
			this.ultraTabPageControl1.Controls.Add(this.SriAutRet);
			this.ultraTabPageControl1.Controls.Add(this.label17);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor16);
			this.ultraTabPageControl1.Controls.Add(this.label16);
			this.ultraTabPageControl1.Controls.Add(this.SriAutFact);
			this.ultraTabPageControl1.Controls.Add(this.label15);
			this.ultraTabPageControl1.Controls.Add(this.SriSerieFact);
			this.ultraTabPageControl1.Controls.Add(this.label14);
			this.ultraTabPageControl1.Controls.Add(this.label11);
			this.ultraTabPageControl1.Controls.Add(this.label10);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor10);
			this.ultraTabPageControl1.Controls.Add(this.spnEmpresa);
			this.ultraTabPageControl1.Controls.Add(this.chkDepura);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(956, 518);
			// 
			// label71
			// 
			this.label71.BackColor = System.Drawing.Color.Transparent;
			this.label71.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label71.Location = new System.Drawing.Point(6, 312);
			this.label71.Name = "label71";
			this.label71.Size = new System.Drawing.Size(944, 16);
			this.label71.TabIndex = 276;
			this.label71.Text = "LIQUIDACIÓN EN COMPRAS";
			this.label71.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label70
			// 
			this.label70.BackColor = System.Drawing.Color.Transparent;
			this.label70.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label70.Location = new System.Drawing.Point(8, 168);
			this.label70.Name = "label70";
			this.label70.Size = new System.Drawing.Size(944, 16);
			this.label70.TabIndex = 275;
			this.label70.Text = "NOTAS DE CRÉDITO";
			this.label70.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label69
			// 
			this.label69.BackColor = System.Drawing.Color.Transparent;
			this.label69.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label69.Location = new System.Drawing.Point(8, 24);
			this.label69.Name = "label69";
			this.label69.Size = new System.Drawing.Size(944, 16);
			this.label69.TabIndex = 274;
			this.label69.Text = "RETENCIONES";
			this.label69.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(40, 267);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(41, 17);
			this.label38.TabIndex = 78;
			this.label38.Text = "Caduca";
			// 
			// ultraDateTimeEditor2
			// 
			this.ultraDateTimeEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CaducaNotaDeCredito"));
			this.ultraDateTimeEditor2.DateTime = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			this.ultraDateTimeEditor2.FormatString = "dd/MMM/yyyy";
			this.ultraDateTimeEditor2.Location = new System.Drawing.Point(144, 264);
			this.ultraDateTimeEditor2.Name = "ultraDateTimeEditor2";
			this.ultraDateTimeEditor2.Size = new System.Drawing.Size(144, 22);
			this.ultraDateTimeEditor2.TabIndex = 77;
			this.ultraDateTimeEditor2.Value = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(40, 235);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(29, 17);
			this.label39.TabIndex = 75;
			this.label39.Text = "Serie";
			// 
			// ultraTextEditor24
			// 
			this.ultraTextEditor24.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SerieNotaDeCredito"));
			this.ultraTextEditor24.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor24.Location = new System.Drawing.Point(144, 232);
			this.ultraTextEditor24.Name = "ultraTextEditor24";
			this.ultraTextEditor24.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor24.TabIndex = 76;
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Location = new System.Drawing.Point(40, 203);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(65, 17);
			this.label40.TabIndex = 73;
			this.label40.Text = "Autorización";
			// 
			// ultraTextEditor25
			// 
			this.ultraTextEditor25.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.AutNotaDeCredito"));
			this.ultraTextEditor25.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor25.Location = new System.Drawing.Point(144, 200);
			this.ultraTextEditor25.Name = "ultraTextEditor25";
			this.ultraTextEditor25.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor25.TabIndex = 74;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(360, 488);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(52, 17);
			this.label36.TabIndex = 72;
			this.label36.Text = "Cant. Emi";
			this.label36.Visible = false;
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(520, 488);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(52, 17);
			this.label35.TabIndex = 71;
			this.label35.Text = "Cant. Min";
			this.label35.Visible = false;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(40, 123);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(41, 17);
			this.label34.TabIndex = 70;
			this.label34.Text = "Caduca";
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(400, 464);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(76, 17);
			this.label33.TabIndex = 69;
			this.label33.Text = "Fecha Emision";
			this.label33.Visible = false;
			// 
			// cmbFechaRcad
			// 
			this.cmbFechaRcad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.FechaRCad"));
			this.cmbFechaRcad.DateTime = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			this.cmbFechaRcad.FormatString = "dd/MMM/yyyy";
			this.cmbFechaRcad.Location = new System.Drawing.Point(144, 120);
			this.cmbFechaRcad.Name = "cmbFechaRcad";
			this.cmbFechaRcad.Size = new System.Drawing.Size(144, 22);
			this.cmbFechaRcad.TabIndex = 68;
			this.cmbFechaRcad.Value = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			// 
			// cmbFechaREmi
			// 
			this.cmbFechaREmi.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.FechaREmi"));
			this.cmbFechaREmi.DateTime = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			this.cmbFechaREmi.FormatString = "dd/MMM/yyyy";
			this.cmbFechaREmi.Location = new System.Drawing.Point(496, 464);
			this.cmbFechaREmi.Name = "cmbFechaREmi";
			this.cmbFechaREmi.Size = new System.Drawing.Size(144, 22);
			this.cmbFechaREmi.TabIndex = 67;
			this.cmbFechaREmi.Value = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			this.cmbFechaREmi.Visible = false;
			// 
			// spnCanRMin
			// 
			this.spnCanRMin.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CantR"));
			this.spnCanRMin.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCanRMin.Location = new System.Drawing.Point(576, 488);
			this.spnCanRMin.MaxValue = 1000000;
			this.spnCanRMin.MinValue = 1;
			this.spnCanRMin.Name = "spnCanRMin";
			this.spnCanRMin.PromptChar = ' ';
			this.spnCanRMin.Size = new System.Drawing.Size(88, 22);
			this.spnCanRMin.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnCanRMin.TabIndex = 66;
			this.spnCanRMin.Value = 1;
			this.spnCanRMin.Visible = false;
			// 
			// spnCanREmi
			// 
			this.spnCanREmi.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.CantR"));
			this.spnCanREmi.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCanREmi.Location = new System.Drawing.Point(424, 488);
			this.spnCanREmi.MaxValue = 1000000;
			this.spnCanREmi.MinValue = 1;
			this.spnCanREmi.Name = "spnCanREmi";
			this.spnCanREmi.PromptChar = ' ';
			this.spnCanREmi.Size = new System.Drawing.Size(88, 22);
			this.spnCanREmi.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnCanREmi.TabIndex = 65;
			this.spnCanREmi.Visible = false;
			// 
			// ultraDateTimeEditor1
			// 
			this.ultraDateTimeEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.FechaCadLiq"));
			this.ultraDateTimeEditor1.DateTime = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			this.ultraDateTimeEditor1.FormatString = "dd/MMM/yyyy";
			this.ultraDateTimeEditor1.Location = new System.Drawing.Point(144, 408);
			this.ultraDateTimeEditor1.Name = "ultraDateTimeEditor1";
			this.ultraDateTimeEditor1.Size = new System.Drawing.Size(144, 22);
			this.ultraDateTimeEditor1.TabIndex = 64;
			this.toolTip1.SetToolTip(this.ultraDateTimeEditor1, "Fecha de Caducidad de Liquidacion de Compra");
			this.ultraDateTimeEditor1.Value = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(40, 347);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(65, 17);
			this.label30.TabIndex = 62;
			this.label30.Text = "Autorización";
			// 
			// ultraTextEditor21
			// 
			this.ultraTextEditor21.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.AutLiquida"));
			this.ultraTextEditor21.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor21.Location = new System.Drawing.Point(144, 344);
			this.ultraTextEditor21.Name = "ultraTextEditor21";
			this.ultraTextEditor21.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor21.TabIndex = 63;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(40, 379);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(29, 17);
			this.label31.TabIndex = 60;
			this.label31.Text = "Serie";
			// 
			// ultraTextEditor22
			// 
			this.ultraTextEditor22.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SerieLiquida"));
			this.ultraTextEditor22.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor22.Location = new System.Drawing.Point(144, 376);
			this.ultraTextEditor22.Name = "ultraTextEditor22";
			this.ultraTextEditor22.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor22.TabIndex = 61;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Location = new System.Drawing.Point(40, 411);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(41, 17);
			this.label32.TabIndex = 59;
			this.label32.Text = "Caduca";
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor3.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.AuditoriaDetalle"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(768, 368);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor3.TabIndex = 58;
			this.ultraCheckEditor3.Text = "Auditoria Detallada";
			this.ultraCheckEditor3.Visible = false;
			// 
			// btSerieRetencion
			// 
			this.btSerieRetencion.Location = new System.Drawing.Point(368, 86);
			this.btSerieRetencion.Name = "btSerieRetencion";
			this.btSerieRetencion.Size = new System.Drawing.Size(24, 26);
			this.btSerieRetencion.TabIndex = 57;
			this.btSerieRetencion.Text = "...";
			this.btSerieRetencion.Visible = false;
			this.btSerieRetencion.Click += new System.EventHandler(this.btSerieRetencion_Click);
			// 
			// btFechaCaduca
			// 
			this.btFechaCaduca.Location = new System.Drawing.Point(920, 480);
			this.btFechaCaduca.Name = "btFechaCaduca";
			this.btFechaCaduca.Size = new System.Drawing.Size(24, 26);
			this.btFechaCaduca.TabIndex = 56;
			this.btFechaCaduca.Text = "...";
			this.btFechaCaduca.Visible = false;
			this.btFechaCaduca.Click += new System.EventHandler(this.btFechaCaduca_Click);
			// 
			// btAutRetencion
			// 
			this.btAutRetencion.Location = new System.Drawing.Point(368, 54);
			this.btAutRetencion.Name = "btAutRetencion";
			this.btAutRetencion.Size = new System.Drawing.Size(24, 26);
			this.btAutRetencion.TabIndex = 55;
			this.btAutRetencion.Text = "...";
			this.btAutRetencion.Visible = false;
			this.btAutRetencion.Click += new System.EventHandler(this.btAutRetencion_Click);
			// 
			// btSerieFactura
			// 
			this.btSerieFactura.Location = new System.Drawing.Point(920, 456);
			this.btSerieFactura.Name = "btSerieFactura";
			this.btSerieFactura.Size = new System.Drawing.Size(24, 26);
			this.btSerieFactura.TabIndex = 54;
			this.btSerieFactura.Text = "...";
			this.btSerieFactura.Visible = false;
			this.btSerieFactura.Click += new System.EventHandler(this.btSerieFactura_Click);
			// 
			// btAutFactura
			// 
			this.btAutFactura.Location = new System.Drawing.Point(920, 424);
			this.btAutFactura.Name = "btAutFactura";
			this.btAutFactura.Size = new System.Drawing.Size(24, 26);
			this.btAutFactura.TabIndex = 53;
			this.btAutFactura.Text = "...";
			this.btAutFactura.Visible = false;
			this.btAutFactura.Click += new System.EventHandler(this.btAutFactura_Click);
			// 
			// cmbFecha
			// 
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SriCaduca"));
			this.cmbFecha.DateTime = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(768, 480);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(144, 22);
			this.cmbFecha.TabIndex = 46;
			this.cmbFecha.Value = new System.DateTime(2018, 5, 7, 0, 0, 0, 0);
			this.cmbFecha.Visible = false;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(40, 91);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(29, 17);
			this.label18.TabIndex = 44;
			this.label18.Text = "Serie";
			// 
			// SriSerieRet
			// 
			this.SriSerieRet.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SriSerieRet"));
			this.SriSerieRet.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.SriSerieRet.Location = new System.Drawing.Point(144, 88);
			this.SriSerieRet.Name = "SriSerieRet";
			this.SriSerieRet.Size = new System.Drawing.Size(144, 22);
			this.SriSerieRet.TabIndex = 45;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(40, 59);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(65, 17);
			this.label19.TabIndex = 42;
			this.label19.Text = "Autorización";
			// 
			// SriAutRet
			// 
			this.SriAutRet.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SriAutRet"));
			this.SriAutRet.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.SriAutRet.Location = new System.Drawing.Point(144, 56);
			this.SriAutRet.Name = "SriAutRet";
			this.SriAutRet.Size = new System.Drawing.Size(144, 22);
			this.SriAutRet.TabIndex = 43;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(672, 400);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(76, 17);
			this.label17.TabIndex = 40;
			this.label17.Text = "Contrib Espec.";
			this.label17.Visible = false;
			// 
			// ultraTextEditor16
			// 
			this.ultraTextEditor16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SriContrib"));
			this.ultraTextEditor16.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor16.Location = new System.Drawing.Point(768, 400);
			this.ultraTextEditor16.Name = "ultraTextEditor16";
			this.ultraTextEditor16.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor16.TabIndex = 41;
			this.ultraTextEditor16.Visible = false;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(672, 424);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(65, 17);
			this.label16.TabIndex = 38;
			this.label16.Text = "Aut. Factura";
			this.label16.Visible = false;
			// 
			// SriAutFact
			// 
			this.SriAutFact.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SriaAutoriz"));
			this.SriAutFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.SriAutFact.Location = new System.Drawing.Point(768, 424);
			this.SriAutFact.Name = "SriAutFact";
			this.SriAutFact.Size = new System.Drawing.Size(144, 22);
			this.SriAutFact.TabIndex = 39;
			this.SriAutFact.Visible = false;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(672, 456);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(70, 17);
			this.label15.TabIndex = 36;
			this.label15.Text = "Serie Factura";
			this.label15.Visible = false;
			// 
			// SriSerieFact
			// 
			this.SriSerieFact.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.SriSerie"));
			this.SriSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.SriSerieFact.Location = new System.Drawing.Point(768, 456);
			this.SriSerieFact.Name = "SriSerieFact";
			this.SriSerieFact.Size = new System.Drawing.Size(144, 22);
			this.SriSerieFact.TabIndex = 37;
			this.SriSerieFact.Visible = false;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(672, 480);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(74, 17);
			this.label14.TabIndex = 34;
			this.label14.Text = "Fecha Caduca";
			this.label14.Visible = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(400, 440);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(85, 17);
			this.label11.TabIndex = 32;
			this.label11.Text = "Aut. SRI Export.";
			this.label11.Visible = false;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(680, 336);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(63, 17);
			this.label10.TabIndex = 31;
			this.label10.Text = "Nº Sucursal";
			this.label10.Visible = false;
			// 
			// ultraTextEditor10
			// 
			this.ultraTextEditor10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.AutSriExp"));
			this.ultraTextEditor10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor10.Location = new System.Drawing.Point(496, 440);
			this.ultraTextEditor10.Name = "ultraTextEditor10";
			this.ultraTextEditor10.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor10.TabIndex = 33;
			this.ultraTextEditor10.Visible = false;
			// 
			// spnEmpresa
			// 
			this.spnEmpresa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.NumEmpresa"));
			this.spnEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnEmpresa.Location = new System.Drawing.Point(768, 336);
			this.spnEmpresa.MaxValue = 20;
			this.spnEmpresa.MinValue = 1;
			this.spnEmpresa.Name = "spnEmpresa";
			this.spnEmpresa.PromptChar = ' ';
			this.spnEmpresa.Size = new System.Drawing.Size(128, 22);
			this.spnEmpresa.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnEmpresa.TabIndex = 30;
			this.spnEmpresa.Value = 1;
			this.spnEmpresa.Visible = false;
			// 
			// chkDepura
			// 
			this.chkDepura.BackColor = System.Drawing.Color.Transparent;
			this.chkDepura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkDepura.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.Auditoria"));
			this.chkDepura.Location = new System.Drawing.Point(680, 368);
			this.chkDepura.Name = "chkDepura";
			this.chkDepura.Size = new System.Drawing.Size(80, 22);
			this.chkDepura.TabIndex = 29;
			this.chkDepura.Text = "Auditoria";
			this.chkDepura.Visible = false;
			this.chkDepura.CheckedChanged += new System.EventHandler(this.ultraCheckEditor2_CheckedChanged);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.ultraCheckEditor10);
			this.ultraTabPageControl3.Controls.Add(this.label27);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor20);
			this.ultraTabPageControl3.Controls.Add(this.label22);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor15);
			this.ultraTabPageControl3.Controls.Add(this.label21);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor14);
			this.ultraTabPageControl3.Controls.Add(this.label20);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor13);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(956, 518);
			// 
			// ultraCheckEditor10
			// 
			this.ultraCheckEditor10.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor10.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor10.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.SSLCorreo"));
			this.ultraCheckEditor10.Location = new System.Drawing.Point(144, 181);
			this.ultraCheckEditor10.Name = "ultraCheckEditor10";
			this.ultraCheckEditor10.Size = new System.Drawing.Size(136, 21);
			this.ultraCheckEditor10.TabIndex = 71;
			this.ultraCheckEditor10.Text = "Habilitar SSL";
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(32, 146);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(37, 17);
			this.label27.TabIndex = 68;
			this.label27.Text = "Puerto";
			// 
			// ultraTextEditor20
			// 
			this.ultraTextEditor20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.PuertoCorreo"));
			this.ultraTextEditor20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor20.Location = new System.Drawing.Point(144, 146);
			this.ultraTextEditor20.Name = "ultraTextEditor20";
			this.ultraTextEditor20.Size = new System.Drawing.Size(248, 22);
			this.ultraTextEditor20.TabIndex = 69;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(32, 112);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(61, 17);
			this.label22.TabIndex = 66;
			this.label22.Text = "Contraseña";
			// 
			// ultraTextEditor15
			// 
			this.ultraTextEditor15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.ClaveCorreo"));
			this.ultraTextEditor15.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor15.Location = new System.Drawing.Point(144, 112);
			this.ultraTextEditor15.Name = "ultraTextEditor15";
			this.ultraTextEditor15.PasswordChar = '*';
			this.ultraTextEditor15.Size = new System.Drawing.Size(248, 22);
			this.ultraTextEditor15.TabIndex = 67;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(32, 78);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(46, 17);
			this.label21.TabIndex = 64;
			this.label21.Text = "Usuario:";
			// 
			// ultraTextEditor14
			// 
			this.ultraTextEditor14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.UsuarioCorreo"));
			this.ultraTextEditor14.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor14.Location = new System.Drawing.Point(144, 78);
			this.ultraTextEditor14.Name = "ultraTextEditor14";
			this.ultraTextEditor14.Size = new System.Drawing.Size(248, 22);
			this.ultraTextEditor14.TabIndex = 65;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(32, 43);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(102, 17);
			this.label20.TabIndex = 62;
			this.label20.Text = "Servidor de Correo:";
			// 
			// ultraTextEditor13
			// 
			this.ultraTextEditor13.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.ServidorCorreo"));
			this.ultraTextEditor13.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor13.Location = new System.Drawing.Point(144, 43);
			this.ultraTextEditor13.Name = "ultraTextEditor13";
			this.ultraTextEditor13.Size = new System.Drawing.Size(248, 22);
			this.ultraTextEditor13.TabIndex = 63;
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.lblProyecto);
			this.ultraTabPageControl4.Controls.Add(this.txtProyecto);
			this.ultraTabPageControl4.Controls.Add(this.lblFiduciario);
			this.ultraTabPageControl4.Controls.Add(this.txtFiduciario);
			this.ultraTabPageControl4.Controls.Add(this.ultraCheckEditor9);
			this.ultraTabPageControl4.Controls.Add(this.ultraLabel2);
			this.ultraTabPageControl4.Controls.Add(this.ultraNumericEditor1);
			this.ultraTabPageControl4.Controls.Add(this.ultraCheckEditor8);
			this.ultraTabPageControl4.Controls.Add(this.ultraCheckEditor5);
			this.ultraTabPageControl4.Controls.Add(this.ultraCheckEditor4);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(956, 518);
			// 
			// lblProyecto
			// 
			this.lblProyecto.AutoSize = true;
			this.lblProyecto.BackColor = System.Drawing.Color.Transparent;
			this.lblProyecto.Location = new System.Drawing.Point(264, 64);
			this.lblProyecto.Name = "lblProyecto";
			this.lblProyecto.Size = new System.Drawing.Size(51, 17);
			this.lblProyecto.TabIndex = 66;
			this.lblProyecto.Text = "Proyecto:";
			this.lblProyecto.Visible = false;
			// 
			// txtProyecto
			// 
			this.txtProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.P_Proyecto"));
			this.txtProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyecto.Location = new System.Drawing.Point(320, 64);
			this.txtProyecto.Name = "txtProyecto";
			this.txtProyecto.Size = new System.Drawing.Size(248, 22);
			this.txtProyecto.TabIndex = 67;
			this.txtProyecto.Visible = false;
			// 
			// lblFiduciario
			// 
			this.lblFiduciario.AutoSize = true;
			this.lblFiduciario.BackColor = System.Drawing.Color.Transparent;
			this.lblFiduciario.Location = new System.Drawing.Point(264, 32);
			this.lblFiduciario.Name = "lblFiduciario";
			this.lblFiduciario.Size = new System.Drawing.Size(56, 17);
			this.lblFiduciario.TabIndex = 64;
			this.lblFiduciario.Text = "Fiduciario:";
			this.lblFiduciario.Visible = false;
			// 
			// txtFiduciario
			// 
			this.txtFiduciario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.P_Fiduciario"));
			this.txtFiduciario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFiduciario.Location = new System.Drawing.Point(320, 32);
			this.txtFiduciario.Name = "txtFiduciario";
			this.txtFiduciario.Size = new System.Drawing.Size(248, 22);
			this.txtFiduciario.TabIndex = 65;
			this.txtFiduciario.Visible = false;
			// 
			// ultraCheckEditor9
			// 
			this.ultraCheckEditor9.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor9.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor9.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.StoreProcedure"));
			this.ultraCheckEditor9.Location = new System.Drawing.Point(48, 129);
			this.ultraCheckEditor9.Name = "ultraCheckEditor9";
			this.ultraCheckEditor9.Size = new System.Drawing.Size(216, 22);
			this.ultraCheckEditor9.TabIndex = 35;
			this.ultraCheckEditor9.Text = "Procedimientos Almacenados Alternos";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(48, 207);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(47, 15);
			this.ultraLabel2.TabIndex = 34;
			this.ultraLabel2.Text = "Usuarios";
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.Location = new System.Drawing.Point(120, 207);
			this.ultraNumericEditor1.MaskInput = "nnn";
			this.ultraNumericEditor1.MaxValue = 999;
			this.ultraNumericEditor1.MinValue = 1;
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(80, 22);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 33;
			// 
			// ultraCheckEditor8
			// 
			this.ultraCheckEditor8.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor8.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor8.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.EliminaPago"));
			this.ultraCheckEditor8.Location = new System.Drawing.Point(48, 95);
			this.ultraCheckEditor8.Name = "ultraCheckEditor8";
			this.ultraCheckEditor8.Size = new System.Drawing.Size(216, 21);
			this.ultraCheckEditor8.TabIndex = 32;
			this.ultraCheckEditor8.Text = "Eliminar Pagos";
			// 
			// ultraCheckEditor5
			// 
			this.ultraCheckEditor5.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor5.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor5.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.EliminaFactura"));
			this.ultraCheckEditor5.Location = new System.Drawing.Point(48, 60);
			this.ultraCheckEditor5.Name = "ultraCheckEditor5";
			this.ultraCheckEditor5.Size = new System.Drawing.Size(216, 22);
			this.ultraCheckEditor5.TabIndex = 31;
			this.ultraCheckEditor5.Text = "Eliminar Facturas";
			// 
			// ultraCheckEditor4
			// 
			this.ultraCheckEditor4.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor4.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor4.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.EliminaAsiento"));
			this.ultraCheckEditor4.Location = new System.Drawing.Point(48, 26);
			this.ultraCheckEditor4.Name = "ultraCheckEditor4";
			this.ultraCheckEditor4.Size = new System.Drawing.Size(208, 21);
			this.ultraCheckEditor4.TabIndex = 30;
			this.ultraCheckEditor4.Text = "Eliminar Asientos";
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.ultraCheckEditor11);
			this.ultraTabPageControl5.Controls.Add(this.ultraLabel1);
			this.ultraTabPageControl5.Controls.Add(this.ultraCheckEditor7);
			this.ultraTabPageControl5.Controls.Add(this.btAnexoTrans);
			this.ultraTabPageControl5.Controls.Add(this.ultraCheckEditor2);
			this.ultraTabPageControl5.Controls.Add(this.chkFormMarzo);
			this.ultraTabPageControl5.Controls.Add(this.ultraCheckEditor6);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(956, 518);
			// 
			// ultraCheckEditor11
			// 
			this.ultraCheckEditor11.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor11.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor11.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.PNatural"));
			this.ultraCheckEditor11.Location = new System.Drawing.Point(32, 129);
			this.ultraCheckEditor11.Name = "ultraCheckEditor11";
			this.ultraCheckEditor11.Size = new System.Drawing.Size(456, 22);
			this.ultraCheckEditor11.TabIndex = 67;
			this.ultraCheckEditor11.Text = "Persona Natural";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraLabel1.Location = new System.Drawing.Point(32, 155);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(472, 43);
			this.ultraLabel1.TabIndex = 66;
			this.ultraLabel1.Text = "Debe estar activada la opción de Anexo transaccional, y solo una de las opciones " +
				"Anexo del 2008 o REOC";
			// 
			// ultraCheckEditor7
			// 
			this.ultraCheckEditor7.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor7.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor7.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.Reoc"));
			this.ultraCheckEditor7.Location = new System.Drawing.Point(32, 48);
			this.ultraCheckEditor7.Name = "ultraCheckEditor7";
			this.ultraCheckEditor7.Size = new System.Drawing.Size(456, 21);
			this.ultraCheckEditor7.TabIndex = 65;
			this.ultraCheckEditor7.Text = "REOC Anexo de compras y Retención de Impuesto a la Renta  por otros Conceptos";
			// 
			// btAnexoTrans
			// 
			this.btAnexoTrans.Location = new System.Drawing.Point(512, 16);
			this.btAnexoTrans.Name = "btAnexoTrans";
			this.btAnexoTrans.Size = new System.Drawing.Size(24, 26);
			this.btAnexoTrans.TabIndex = 64;
			this.btAnexoTrans.Text = "...";
			this.btAnexoTrans.Visible = false;
			// 
			// ultraCheckEditor2
			// 
			this.ultraCheckEditor2.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.At_Fecha_Cont"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(40, 224);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(488, 22);
			this.ultraCheckEditor2.TabIndex = 63;
			this.ultraCheckEditor2.Text = "Anexo Transaccional por Fecha de Contabilización. desarrollo personalizado NO USA" +
				"R";
			// 
			// chkFormMarzo
			// 
			this.chkFormMarzo.BackColor = System.Drawing.Color.Transparent;
			this.chkFormMarzo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkFormMarzo.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.FormMarzo"));
			this.chkFormMarzo.Location = new System.Drawing.Point(376, 16);
			this.chkFormMarzo.Name = "chkFormMarzo";
			this.chkFormMarzo.Size = new System.Drawing.Size(136, 22);
			this.chkFormMarzo.TabIndex = 62;
			this.chkFormMarzo.Text = "Anexo Transaccional";
			this.chkFormMarzo.Visible = false;
			// 
			// ultraCheckEditor6
			// 
			this.ultraCheckEditor6.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor6.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor6.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.Anexo2008"));
			this.ultraCheckEditor6.Location = new System.Drawing.Point(32, 24);
			this.ultraCheckEditor6.Name = "ultraCheckEditor6";
			this.ultraCheckEditor6.Size = new System.Drawing.Size(272, 22);
			this.ultraCheckEditor6.TabIndex = 61;
			this.ultraCheckEditor6.Text = "Anexo Transaccional Simplificado del 2008";
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.label29);
			this.ultraTabPageControl6.Controls.Add(this.txtInteres2);
			this.ultraTabPageControl6.Controls.Add(this.label28);
			this.ultraTabPageControl6.Controls.Add(this.txtInteres1);
			this.ultraTabPageControl6.Controls.Add(this.chkAutoImpresor);
			this.ultraTabPageControl6.Controls.Add(this.label25);
			this.ultraTabPageControl6.Controls.Add(this.ultraTextEditor19);
			this.ultraTabPageControl6.Controls.Add(this.label24);
			this.ultraTabPageControl6.Controls.Add(this.ultraTextEditor18);
			this.ultraTabPageControl6.Controls.Add(this.label23);
			this.ultraTabPageControl6.Controls.Add(this.ultraTextEditor17);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(956, 518);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(32, 190);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(49, 17);
			this.label29.TabIndex = 67;
			this.label29.Text = "Interes 2";
			// 
			// txtInteres2
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInteres2.Appearance = appearance47;
			this.txtInteres2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Interes2"));
			this.txtInteres2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteres2.FormatString = "#,##0.00";
			this.txtInteres2.Location = new System.Drawing.Point(192, 190);
			this.txtInteres2.Name = "txtInteres2";
			this.txtInteres2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteres2.PromptChar = ' ';
			this.txtInteres2.Size = new System.Drawing.Size(152, 22);
			this.txtInteres2.TabIndex = 66;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(32, 155);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(49, 17);
			this.label28.TabIndex = 65;
			this.label28.Text = "Interes 1";
			// 
			// txtInteres1
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInteres1.Appearance = appearance48;
			this.txtInteres1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Interes1"));
			this.txtInteres1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteres1.FormatString = "#,##0.00";
			this.txtInteres1.Location = new System.Drawing.Point(192, 155);
			this.txtInteres1.Name = "txtInteres1";
			this.txtInteres1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteres1.PromptChar = ' ';
			this.txtInteres1.Size = new System.Drawing.Size(152, 22);
			this.txtInteres1.TabIndex = 64;
			// 
			// chkAutoImpresor
			// 
			this.chkAutoImpresor.BackColor = System.Drawing.Color.Transparent;
			this.chkAutoImpresor.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkAutoImpresor.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteo, "Seteo.AutoImpresor"));
			this.chkAutoImpresor.Location = new System.Drawing.Point(192, 121);
			this.chkAutoImpresor.Name = "chkAutoImpresor";
			this.chkAutoImpresor.Size = new System.Drawing.Size(136, 21);
			this.chkAutoImpresor.TabIndex = 63;
			this.chkAutoImpresor.Text = "AutoImpresor";
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(32, 86);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(57, 17);
			this.label25.TabIndex = 44;
			this.label25.Text = "Secuencial";
			// 
			// ultraTextEditor19
			// 
			this.ultraTextEditor19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.AutoImpSecInicio"));
			this.ultraTextEditor19.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor19.Location = new System.Drawing.Point(192, 86);
			this.ultraTextEditor19.Name = "ultraTextEditor19";
			this.ultraTextEditor19.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor19.TabIndex = 45;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(32, 52);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(152, 17);
			this.label24.TabIndex = 42;
			this.label24.Text = "Establecimiento / Pto emisión";
			// 
			// ultraTextEditor18
			// 
			this.ultraTextEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.AutoImpEstablec"));
			this.ultraTextEditor18.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor18.Location = new System.Drawing.Point(192, 52);
			this.ultraTextEditor18.Name = "ultraTextEditor18";
			this.ultraTextEditor18.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor18.TabIndex = 43;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(32, 17);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(124, 17);
			this.label23.TabIndex = 40;
			this.label23.Text = "Número de Autorización";
			// 
			// ultraTextEditor17
			// 
			this.ultraTextEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.AutoImpAutoriza"));
			this.ultraTextEditor17.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor17.Location = new System.Drawing.Point(192, 17);
			this.ultraTextEditor17.Name = "ultraTextEditor17";
			this.ultraTextEditor17.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor17.TabIndex = 41;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 560);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 18;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(88, 560);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 19;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl5);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl6);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl7);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl8);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl9);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl10);
			this.ultraTabControl1.Controls.Add(this.Documentos);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl11);
			this.ultraTabControl1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraTabControl1.Location = new System.Drawing.Point(8, 8);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(960, 544);
			this.ultraTabControl1.TabIndex = 32;
			ultraTab1.TabPage = this.ultraTabPageControl2;
			ultraTab1.Text = "General";
			ultraTab2.TabPage = this.ultraTabPageControl7;
			ultraTab2.Text = "Representante Legal";
			ultraTab3.TabPage = this.ultraTabPageControl8;
			ultraTab3.Text = "Socios";
			ultraTab4.TabPage = this.ultraTabPageControl9;
			ultraTab4.Text = "Contactos";
			ultraTab5.TabPage = this.ultraTabPageControl10;
			ultraTab5.Text = "Cumplimientos";
			ultraTab6.TabPage = this.Documentos;
			ultraTab6.Text = "Documentos";
			ultraTab7.TabPage = this.ultraTabPageControl11;
			ultraTab7.Text = "Recordatorios";
			ultraTab8.TabPage = this.ultraTabPageControl1;
			ultraTab8.Text = "Complemento";
			ultraTab9.TabPage = this.ultraTabPageControl3;
			ultraTab9.Text = "Correo Electrónico";
			ultraTab9.Visible = false;
			ultraTab10.TabPage = this.ultraTabPageControl4;
			ultraTab10.Text = "Seguridad";
			ultraTab10.Visible = false;
			ultraTab11.TabPage = this.ultraTabPageControl5;
			ultraTab11.Text = "Anexo transaccional";
			ultraTab11.Visible = false;
			ultraTab12.TabPage = this.ultraTabPageControl6;
			ultraTab12.Text = "AutoImpresor";
			ultraTab12.Visible = false;
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
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
																																															 ultraTab11,
																																															 ultraTab12});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(956, 518);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// Seteo
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(976, 590);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "Seteo";
			this.Text = "Datos de la Empresa";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Seteo_Closing);
			this.Load += new System.EventHandler(this.Seteo_Load);
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtActividadEconomica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicioActividades)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaInscripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaResolucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNResolucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAgenteDeRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor23)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLicenciaContador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadorNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbNENacionalidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferenciaRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInterseccionRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParroquiaRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelularRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMailRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConvencionalRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrincipalRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.CedPasRep)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacionRL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepresentanteRL)).EndInit();
			this.ultraTabPageControl8.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridSocios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSocios)).EndInit();
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridContactos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContactos)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridCumplimientos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodoMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObligaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresasPublicas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.Documentos.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			this.ultraTabPageControl11.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRecordatorios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRecordatorios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraDateTimeEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor24)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor25)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaRcad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaREmi)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCanRMin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCanREmi)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDateTimeEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor22)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.SriSerieRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.SriAutRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.SriAutFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.SriSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEmpresa)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor13)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFiduciario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtInteres2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsSeteo.Update();
			this.cdsSocios.Update();
			this.cdsContactos.Update();
			
			#region Guarda Cumplimientos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCumplimientos.Rows.All)
			{
				string sSQL = string.Format("Exec SeteoGuardaCumplimientos {0}, {1}, {2}, {3}, '{4}', '{5}'", 
					(int)dr.Cells["idCumplimiento"].Value, (bool)dr.Cells["Activo"].Value, (int)dr.Cells["idEmpresasPublicas"].Value, (int)dr.Cells["IdPlanificacion"].Value,
					dr.Cells["Documento"].Value.ToString(), dr.Cells["Descripcion"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteo, sSQL, true);
			}
			#endregion Guarda Cumplimientos

			#region Guarda Documentos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				DateTime dtVigencia = DateTime.Parse("01/01/2000");
				if (dr.Cells["FechaVigencia"].Value != System.DBNull.Value) dtVigencia = Convert.ToDateTime(dr.Cells["FechaVigencia"].Value);
				DateTime dtFechaConstitucion = DateTime.Parse("01/01/2000");
				if (dr.Cells["FechaConstitucion"].Value != System.DBNull.Value) dtFechaConstitucion = Convert.ToDateTime(dr.Cells["FechaConstitucion"].Value);
				DateTime dtFechaInscripcion = DateTime.Parse("01/01/2000");
				if (dr.Cells["FechaInscripcion"].Value != System.DBNull.Value) dtFechaInscripcion = Convert.ToDateTime(dr.Cells["FechaInscripcion"].Value);
				DateTime dtFechaActualizacion = DateTime.Parse("01/01/2000");
				if (dr.Cells["FechaActualizacion"].Value != System.DBNull.Value) dtFechaActualizacion = Convert.ToDateTime(dr.Cells["FechaActualizacion"].Value);

				string sSQL = string.Format("Exec SeteoGuardaDocumentos {0}, {1}, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}'", 
					(int)dr.Cells["idDocumento"].Value, (bool)dr.Cells["Activo"].Value, dr.Cells["Documento"].Value.ToString(), dr.Cells["Descripcion"].Value.ToString(),
					dtVigencia.ToString("yyyyMMdd"), dtFechaConstitucion.ToString("yyyyMMdd"), dtFechaInscripcion.ToString("yyyyMMdd"), dtFechaActualizacion.ToString("yyyyMMdd"), 
					dr.Cells["Ubicacion"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteo, sSQL, true);
			}					
			#endregion Guarda Documentos
			
			#region Guarda Recordatorios
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRecordatorios.Rows.All)
			{
				string sSQL = string.Format("Exec SeteoActualizaRecordatorios {0}, {1}, '{2}', '{3}'", 
					(int)dr.Cells["idRecordatorio"].Value, (bool)dr.Cells["Estado"].Value, dr.Cells["Notas"].Value.ToString(), dr.Cells["Ubicacion"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteo, sSQL, true);
			}
			#endregion Guarda Recordatorios
			//			// Actualiza Clave Principal
			//			string stSelect = "EXEC InicioIDP " + this.spnEmpresa.Value.ToString();
			//			Funcion.EjecutaSQL(cdsSeteo, stSelect);

			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void Seteo_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 7, 'Prop. Gen.'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			if (chkAutoImpresor.Checked) chkAutoImpresor.Enabled = false;

			this.cmbTipoEmpresa.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select idTipoEmpresa, Descripcion From TipoEmpresa");
			this.cmbAgenteDeRetencion.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select idAgenteRetencion, Nombre From AgenteRetencion");
			this.cmbNEProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec Listas 0, 0");
			this.cmbNENacionalidad.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec Listas 11, 0");
			this.cmbProvinciaRL.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec Listas 0, 0");

			this.cmbEmpresasPublicas.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select idEmpresasPublicas, Descripcion From EmpresasPublicas");
			this.cmbPeriodoMes.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select IdPlanificacion, Descripcion From Planificacion");

			this.dtFechaRecordatorios.Value = DateTime.Today;

			if (Funcion.bEjecutaSQL(cdsSeteo, "Exec SeteoGExiste 'FTRPCO'")) 				
			{
				txtFiduciario.Visible=true;
				lblFiduciario.Visible=true;
				txtProyecto.Visible=true;
				lblProyecto.Visible=true;
			}
		}

		private void CedPasRep_ValueChanged(object sender, System.EventArgs e)
		{
		}

		private void CedPasRep_Validated(object sender, System.EventArgs e)
		{
			if (CedPasRep.Text != "P" && CedPasRep.Text != "C")
			{
				CedPasRep.Text = "C";
				CedPasRep.Select();
				MessageBox.Show("Solo se Admite C para Cedula y P para Pasaporte", "Información");
			}		
		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(SriSerieFact, "");
			errorProvider.SetError(SriSerieRet, "");
			errorProvider.SetError(SriAutRet, "");
			errorProvider.SetError(SriAutFact, "");
			errorProvider.SetError(cmbFecha, "");
			if (cmbFecha.Value == null || cmbFecha.Value == System.DBNull.Value)
			{
				errorProvider.SetError(cmbFecha, "Ingrese Fecha");
				bOk = false;
			}
			return bOk;
		}

		private void btAutFactura_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtCaduca = (DateTime) cmbFecha.Value;
			string stAutFact = SriAutFact.Text.Trim();
			if (stAutFact.Length == 0)
			{
				errorProvider.SetError(SriAutFact, "Ingrese Autorización de Factura");
				return;
			}
	
			if (DialogResult.No == MessageBox.Show("Desea Reemplazar todas las Ventas "
				+ "de Fechas anteriores a: " + dtCaduca.ToString("dd/MMM/yyyy") 
				+ " y que no tengan autorización con esta Autorización: " 
        + stAutFact, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			string stSelect = "Update Compra Set AutFactura = '" + stAutFact
				+ "' Where idTipoFactura = 1 "
				+ "AND LEN(LTRIM(RTRIM(IsNull(Compra.AutFactura, '')))) = 0 AND Compra.Fecha < '"
				+ dtCaduca.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
      MessageBox.Show("Reemplazado");
		}

		private void btSerieFactura_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtCaduca = (DateTime) cmbFecha.Value;
			string stAutFact = SriSerieFact.Text.Trim();
			if (stAutFact.Length == 0)
			{
				errorProvider.SetError(SriSerieFact, "Ingrese Serie de Factura");
				return;
			}
	
			if (DialogResult.No == MessageBox.Show("Desea Reemplazar todas las Ventas "
				+ "de Fechas anteriores a: " + dtCaduca.ToString("dd/MMM/yyyy") 
				+ " y que no tengan Serie con esta Serie: " 
				+ stAutFact, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			string stSelect = "Update Compra Set SerieFactura = '" + stAutFact
				+ "' Where idTipoFactura = 1 "
				+ "AND LEN(LTRIM(RTRIM(IsNull(Compra.SerieFactura, '')))) = 0 AND Compra.Fecha < '"
				+ dtCaduca.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
			MessageBox.Show("Reemplazado");
		}

		private void btAutRetencion_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtCaduca = (DateTime) cmbFecha.Value;
			string stAutFact = SriAutRet.Text.Trim();
			if (stAutFact.Length == 0)
			{
				errorProvider.SetError(SriAutRet, "Ingrese Autorizaciòn de Retención");
				return;
			}
	
			if (DialogResult.No == MessageBox.Show("Desea Reemplazar todas las Retenciones "
				+ "de Fechas anteriores a: " + dtCaduca.ToString("dd/MMM/yyyy") 
				+ " y que no tengan Autorizaciòn con esta Autorización: " 
				+ stAutFact, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			string stSelect = "Update Retencion Set AutorizaRet = '"+ stAutFact
				+ "' Where IngresoEgreso = 2 "
				+ "AND LEN(LTRIM(RTRIM(IsNull(Retencion.AutorizaRet, '')))) = 0 "
				+ "AND Retencion.Fecha < '"
				+ dtCaduca.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
			MessageBox.Show("Reemplazado", "Información");		
		}

		private void btSerieRetencion_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtCaduca = (DateTime) cmbFecha.Value;
			string stAutFact = SriSerieRet.Text.Trim();
			if (stAutFact.Length == 0)
			{
				errorProvider.SetError(SriSerieRet, "Ingrese Serie de Retención");
				return;
			}
	
			if (DialogResult.No == MessageBox.Show("Desea Reemplazar todas las Retenciones "
				+ "de Fechas anteriores a: " + dtCaduca.ToString("dd/MMM/yyyy") 
				+ " y que no tengan Serie con esta Serie: " 
				+ stAutFact, "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			string stSelect = "Update Retencion Set SerieRet = '" + stAutFact
				+ "' Where IngresoEgreso = 2 "
				+ "AND LEN(LTRIM(RTRIM(IsNull(Retencion.SerieRet, '')))) = 0 AND Retencion.Fecha < '"
				+ dtCaduca.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
			MessageBox.Show("Reemplazado", "Información");		
		}

		private void btFechaCaduca_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtCaduca = (DateTime) cmbFecha.Value;
	
			if (DialogResult.No == MessageBox.Show("Desea Reemplazar todas las Ventas "
				+ "de Fechas anteriores a: " + dtCaduca.ToString("dd/MMM/yyyy") 
				+ " y que no tengan Caducidad con esta Caducidad: " 
				+ dtCaduca.ToString("dd/MMM/yyyy"), "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			string stSelect = "Update Compra Set FechaCaducidad = '" 
				+ dtCaduca.ToString("yyyyMMdd") + "' Where idTipoFactura = 1 "
				+ "AND Compra.Fecha < '" + dtCaduca.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
			MessageBox.Show("Reemplazado", "Información");		
		}

		private void btAnexoTrans_Click(object sender, System.EventArgs e)
		{
			ActualizaRetencion miActualiza = new ActualizaRetencion();
			miActualiza.MdiParent = this.MdiParent;
			miActualiza.Show();
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void ultraCheckEditor2_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!chkDepura.Checked)
			{
				using (CompraAutoriza miAutoriza = new CompraAutoriza())
				{
					if (DialogResult.OK != miAutoriza.ShowDialog())
					{
						chkDepura.Select();
						chkDepura.Refresh();
						chkDepura.Checked = true;
						chkDepura.Refresh();
						chkDepura.Select();
						chkDepura.Update();
					}
					else
					{
						string stSelect = "INSERT INTO AUDITORIA (IDTIPOAUDITA, IDAUDITAESTADO, "
							+ "CODIGO, IDUNICA, CAMPO, VALOR)	VALUES (42, 2, 'Autoriza', 1, 'Audita', 0)";
						Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
					}
				}
			}
		}

		private void Seteo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Prop. Gen.'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
		}
		
		private void cmbNEProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNEProvincia.ActiveRow != null)
				this.cmbNECanton.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec Listas 1, {0}", Convert.ToInt32(this.cmbNEProvincia.Value)));
		}

		private void cmbNECanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNECanton.ActiveRow != null)
				this.cmbNEParroquia.DataSource = Funcion.dvProcedimiento (cdsSeteo, string.Format("Exec Listas 2, {0}", Convert.ToInt32(this.cmbNECanton.Value)));
		}

		private void cmbNEParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNEParroquia.ActiveRow != null)
				this.cmbNEBarrio.DataSource = Funcion.dvProcedimiento (cdsSeteo, string.Format("Exec Listas 3, {0}", Convert.ToInt32(this.cmbNEParroquia.Value)));
		}
		
		private void cmbProvinciaRL_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbProvinciaRL.ActiveRow != null)
				this.cmbCantonRL.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec Listas 1, {0}", Convert.ToInt32(this.cmbProvinciaRL.Value)));
		}

		private void cmbCantonRL_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCantonRL.ActiveRow != null)
				this.cmbParroquiaRL.DataSource = Funcion.dvProcedimiento (cdsSeteo, string.Format("Exec Listas 2, {0}", Convert.ToInt32(this.cmbCantonRL.Value)));
		}

		private void cdsSocios_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSocios.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsContactos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsContactos.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
		
		private void ultraGrid3_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnRecordatorios")
			{
				frmRecordatorios Recordatorio = new frmRecordatorios((int)e.Cell.Row.Cells["idCumplimiento"].Value, e.Cell.Row.Cells["Documento"].Value.ToString(), 1);							
				Recordatorio.ShowDialog();									
			}
		}

		private void ultraGrid3_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoEmpresa_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipoEmpresa.ActiveRow != null)
			{
				this.uGridCumplimientos.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec SeteoListaCumplimientos {0}, 1", (int)this.cmbTipoEmpresa.Value));
				this.uGridDocumentos.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec SeteoListaCumplimientos {0}, 2", (int)this.cmbTipoEmpresa.Value));
			}
		}
		
		private void uGridCumplimientos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridCumplimientos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnFechas")
			{
				frmSeteoFechasCumplimientos SFC = new frmSeteoFechasCumplimientos((int)e.Cell.Row.Cells["idCumplimiento"].Value, (int)e.Cell.Row.Cells["IdPlanificacion"].Value);
				SFC.ShowDialog();
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec SeteoConsultaRecordatorios '{0}', {1}", Convert.ToDateTime(this.dtFechaRecordatorios.Value).ToString("yyyyMMdd"), Convert.ToBoolean(this.optFiltro.Value));
			this.uGridRecordatorios.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void btnUsuarios_Click(object sender, System.EventArgs e)
		{
			frmUsuariosCumplimiento UC = new frmUsuariosCumplimiento();
			UC.ShowDialog();
		}

		private void uGridCumplimientos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridCumplimientos);
		}

		private void uGridSocios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSocios);
		}

		private void uGridContactos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridContactos);
		}

		private void uGridDocumentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDocumentos);
		}

		private void uGridRecordatorios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridRecordatorios);
		}

		private void uGridDocumentos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridDocumentos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnSubir")
			{
				try
				{
					string sArchivo = e.Cell.Row.Cells["Ubicacion"].Value.ToString();
					if (sArchivo.Length > 0)
					{
						if (File.Exists(sArchivo))
						{							
							if (DialogResult.Yes == MessageBox.Show(string.Format("El archivo del documento {0} ya fue subido\n\n¿Desea Visualizarlo?", e.Cell.Row.Cells["Documento"].Value.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
							{
								Process.Start(sArchivo, "");
								return;
							}
							else return;
						}
					}

					OpenFileDialog oArchivo = new OpenFileDialog();

					oArchivo.Filter = "Archivos pdf (*.pdf)|*.pdf";
					oArchivo.AddExtension = true;
					
					if (oArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && oArchivo.FileName.Length > 0)
					{						
						e.Cell.Row.Cells["Ubicacion"].Value = @"\\192.168.1.3\doclatinium\" + MenuLatinium.stDirFacturacion + @"\" + e.Cell.Row.Cells["Documento"].Value.ToString() + ".pdf";
						File.Copy(oArchivo.FileName, e.Cell.Row.Cells["Ubicacion"].Value.ToString());
							
						MessageBox.Show("Importación Finalizada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								
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

			if (e.Cell.Column.ToString() == "btnVer")
			{
				string sArchivo = e.Cell.Row.Cells["Ubicacion"].Value.ToString();
				
				if (!File.Exists(sArchivo))
				{	
					MessageBox.Show("No existe archivo de este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (File.Exists(sArchivo))					
					Process.Start(sArchivo, "");
			}
		}

		private void uGridRecordatorios_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnSubir")
			{
				try
				{
					string sArchivo = e.Cell.Row.Cells["Ubicacion"].Value.ToString();
					if (sArchivo.Length > 0)
					{
						if (File.Exists(sArchivo))
						{							
							if (DialogResult.Yes == MessageBox.Show(string.Format("El archivo del documento {0} ya fue subido\n\n¿Desea Visualizarlo?", e.Cell.Row.Cells["Documento"].Value.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
							{
								Process.Start(sArchivo, "");
								return;
							}
							else return;
						}
					}

					OpenFileDialog oArchivo = new OpenFileDialog();

					oArchivo.Filter = "Archivos pdf (*.pdf)|*.pdf";
					oArchivo.AddExtension = true;
					
					if (oArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && oArchivo.FileName.Length > 0)
					{						
						e.Cell.Row.Cells["Ubicacion"].Value = @"\\192.168.1.3\doclatinium\" + MenuLatinium.stDirFacturacion + @"\" + e.Cell.Row.Cells["Documento"].Value.ToString() + ".pdf";
						File.Copy(oArchivo.FileName, e.Cell.Row.Cells["Ubicacion"].Value.ToString());
							
						MessageBox.Show("Importación Finalizada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						
						e.Cell.Row.Cells["Estado"].Value = true;
		
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

			if (e.Cell.Column.ToString() == "btnVer")
			{
				string sArchivo = e.Cell.Row.Cells["Ubicacion"].Value.ToString();
				
				if (!File.Exists(sArchivo))
				{	
					MessageBox.Show("No existe archivo de este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (File.Exists(sArchivo))					
					Process.Start(sArchivo, "");
			}
		}

		private void uGridRecordatorios_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
