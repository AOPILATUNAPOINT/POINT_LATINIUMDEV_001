using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolPersonal.
	/// </summary>
	public class RolPersonal : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optSistema;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAfiliacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoSectorial;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraPictureBox picFoto;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private C1.Data.C1DataSet cdsSucursal;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtSalida;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtNacimiento;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIngreso;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtNovedad;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkIess;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFondoReserva;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSueldo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtJornal;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRenta;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.Misc.UltraButton btDepartamento;
		private Infragistics.Win.Misc.UltraButton btCargo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVendedor;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCobrador;
		private C1.Data.C1DataSet cdsRolConcepto;
		private C1.Data.C1DataSet cdsCiudad;
		private C1.Data.C1DataSet cdsProvincia;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo1;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCargas;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optSexo;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet ultraOptionSet1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private System.Windows.Forms.Label label27;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor6;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor7;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor8;
		private System.Windows.Forms.Label label32;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmail;
		private System.Windows.Forms.Label label33;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label label34;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvCuenta;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor4;
		private System.Windows.Forms.Label label37;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor5;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl8;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbNivel;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTitulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPersonal;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbParentesco;
		private Infragistics.Win.Misc.UltraButton btNivel;
		private Infragistics.Win.Misc.UltraButton btTitulo;
		private Infragistics.Win.Misc.UltraButton btParentesco;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsAsientoDistrib;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor9;
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbCedula;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor6;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.Misc.UltraButton btCargasF;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCalendario;
		private Infragistics.Win.Misc.UltraButton btHorario;
		private Infragistics.Win.Misc.UltraButton btDistribucion;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.Misc.UltraButton btHorario1;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo1;
		private Infragistics.Win.Misc.UltraButton btHoraExtra;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton btCiudad;
		private Infragistics.Win.Misc.UltraButton btProvincia;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private C1.Data.C1DataSet cdsBanco;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor7;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor8;
		private System.Windows.Forms.Label label43;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor10;
		private System.Windows.Forms.Label label44;
		private System.Windows.Forms.Label label45;
		private C1.Data.C1DataSet cdsRelacionT;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo4;
		private C1.Data.C1DataView cdvSucursal;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl11;
		private Infragistics.Win.Misc.UltraButton btTotalGasto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGasto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor9;
		private System.Windows.Forms.Label label51;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor10;
		private System.Windows.Forms.Label label52;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl12;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid5;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesdeV;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHastaV;
		private Infragistics.Win.Misc.UltraButton btRepVacacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor11;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor12;
		private System.Windows.Forms.Label label54;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor13;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor14;
		private System.Windows.Forms.Label label56;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor15;
		private System.Windows.Forms.Label label57;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor16;
		private System.Windows.Forms.Label label58;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor17;
		private System.Windows.Forms.Label label59;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor18;
		private System.Windows.Forms.Label label60;
		private System.Windows.Forms.Label label61;
		private System.Windows.Forms.Label label62;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo5;
		private System.Windows.Forms.Label label63;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRegSalida;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRegIngreso;
		private C1.Data.C1DataSet cdsTipoContrato;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor11;
		private System.Windows.Forms.Label label64;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFiltro;
		private System.Windows.Forms.Label lblFiltro;
		private Infragistics.Win.Misc.UltraButton btnReingreso;
		private C1.Data.C1DataSet cdsRolReingreso;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdReingreso;
		private System.Windows.Forms.Label label65;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSupervisor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSupervisor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDistrib;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbJefes;
		private System.Windows.Forms.Label label66;
		private Infragistics.Win.Misc.UltraButton btAdicionales;
		private System.Windows.Forms.Label label67;
		private System.Windows.Forms.Label label68;
		private System.Windows.Forms.Label label69;
		private System.Windows.Forms.Label label70;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDiscapacidad;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentajeDiscapacidad;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optVivienda;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVehiculo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTipoDiscapacidad;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumConadis;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVestimenta;
		private System.Windows.Forms.Label label50;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEducacion;
		private System.Windows.Forms.Label label49;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSalud;
		private System.Windows.Forms.Label label48;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAlimentacion;
		private System.Windows.Forms.Label label47;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVivienda;
		private System.Windows.Forms.Label label46;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo2;
		private System.Windows.Forms.Label lblPersonalDiscapacitado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor4;
		private System.ComponentModel.IContainer components;

		public RolPersonal()
		{
			InitializeComponent();
		}
		int IdPersonal = 0;
		public RolPersonal(int idPersonal)
		{
			InitializeComponent();
			IdPersonal = idPersonal;
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteCiudad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Provincia", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteCiudad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolReingreso", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReingresos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolTipoContrato", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasAlerta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolRelacionTrabajo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Relacion");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolCargo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolDepartamento", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPais");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolHorario", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesHExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesHExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesHExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesHExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesHExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoHExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoHExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoDoble");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AsientoDistGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribucion");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PersTitulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersTitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Titulo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PersNivel", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nivel");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PersEstudios", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstudios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivel", -1, "cmbNivel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Institucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTitulo", -1, "cmbTitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaHasta");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PersParentesco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carga");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PersFamiliar", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFamiliar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco", -1, "cmbParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaNacimiento");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolVacacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolVacacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Horas");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolSueldo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolObservacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolObservacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolUniforme", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolUniforme");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Uniforme");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(RolPersonal));
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbCedula = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.spnCargas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.picFoto = new Infragistics.Win.UltraWinEditors.UltraPictureBox();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCodigoSectorial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.optSistema = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label5 = new System.Windows.Forms.Label();
			this.txtAfiliacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lblFiltro = new System.Windows.Forms.Label();
			this.cmbFiltro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCiudad = new C1.Data.C1DataSet();
			this.ultraTextEditor11 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label64 = new System.Windows.Forms.Label();
			this.ultraNumericEditor8 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label43 = new System.Windows.Forms.Label();
			this.ultraNumericEditor7 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label42 = new System.Windows.Forms.Label();
			this.btProvincia = new Infragistics.Win.Misc.UltraButton();
			this.btCiudad = new Infragistics.Win.Misc.UltraButton();
			this.ultraTextEditor9 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label40 = new System.Windows.Forms.Label();
			this.optSexo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsProvincia = new C1.Data.C1DataSet();
			this.ultraCombo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label12 = new System.Windows.Forms.Label();
			this.ultraCalendarCombo2 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.grdReingreso = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsRolReingreso = new C1.Data.C1DataSet();
			this.btnReingreso = new Infragistics.Win.Misc.UltraButton();
			this.label63 = new System.Windows.Forms.Label();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.ultraCombo5 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsTipoContrato = new C1.Data.C1DataSet();
			this.label62 = new System.Windows.Forms.Label();
			this.txtRegSalida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label61 = new System.Windows.Forms.Label();
			this.txtRegIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraCombo4 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsRelacionT = new C1.Data.C1DataSet();
			this.label45 = new System.Windows.Forms.Label();
			this.label44 = new System.Windows.Forms.Label();
			this.ultraTextEditor10 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label41 = new System.Windows.Forms.Label();
			this.ultraCalendarCombo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label15 = new System.Windows.Forms.Label();
			this.dtNovedad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.dtSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label10 = new System.Windows.Forms.Label();
			this.dtNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label9 = new System.Windows.Forms.Label();
			this.dtIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbJefes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label66 = new System.Windows.Forms.Label();
			this.chkSupervisor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbSupervisor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label65 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraNumericEditor6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.label34 = new System.Windows.Forms.Label();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.label33 = new System.Windows.Forms.Label();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label27 = new System.Windows.Forms.Label();
			this.chkCobrador = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkVendedor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btCargo = new Infragistics.Win.Misc.UltraButton();
			this.btDepartamento = new Infragistics.Win.Misc.UltraButton();
			this.txtJornal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtSueldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.chkFondoReserva = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkRenta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkIess = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSucursal = new C1.Data.C1DataSet();
			this.label14 = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label13 = new System.Windows.Forms.Label();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.btHoraExtra = new Infragistics.Win.Misc.UltraButton();
			this.btHorario1 = new Infragistics.Win.Misc.UltraButton();
			this.btDistribucion = new Infragistics.Win.Misc.UltraButton();
			this.btHorario = new Infragistics.Win.Misc.UltraButton();
			this.cmbCalendario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label39 = new System.Windows.Forms.Label();
			this.cmbDistrib = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsAsientoDistrib = new C1.Data.C1DataSet();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label25 = new System.Windows.Forms.Label();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.ultraOptionSet1 = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtNumConadis = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTipoDiscapacidad = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkVehiculo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label70 = new System.Windows.Forms.Label();
			this.optVivienda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label69 = new System.Windows.Forms.Label();
			this.txtPorcentajeDiscapacidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkDiscapacidad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label67 = new System.Windows.Forms.Label();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraNumericEditor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label37 = new System.Windows.Forms.Label();
			this.ultraNumericEditor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label38 = new System.Windows.Forms.Label();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label36 = new System.Windows.Forms.Label();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label35 = new System.Windows.Forms.Label();
			this.label32 = new System.Windows.Forms.Label();
			this.txtEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.ultraTextEditor8 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.ultraTextEditor7 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label29 = new System.Windows.Forms.Label();
			this.ultraTextEditor6 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label28 = new System.Windows.Forms.Label();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label68 = new System.Windows.Forms.Label();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbTitulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbNivel = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl8 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbParentesco = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btRepVacacion = new Infragistics.Win.Misc.UltraButton();
			this.cmbHastaV = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbDesdeV = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid4 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl11 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraCheckEditor4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraNumericEditor14 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label56 = new System.Windows.Forms.Label();
			this.ultraNumericEditor11 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label53 = new System.Windows.Forms.Label();
			this.ultraNumericEditor12 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label54 = new System.Windows.Forms.Label();
			this.ultraNumericEditor13 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label55 = new System.Windows.Forms.Label();
			this.ultraNumericEditor10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label52 = new System.Windows.Forms.Label();
			this.ultraNumericEditor9 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label51 = new System.Windows.Forms.Label();
			this.btTotalGasto = new Infragistics.Win.Misc.UltraButton();
			this.txtTotalGasto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraNumericEditor18 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label60 = new System.Windows.Forms.Label();
			this.ultraNumericEditor15 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label57 = new System.Windows.Forms.Label();
			this.ultraNumericEditor16 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label58 = new System.Windows.Forms.Label();
			this.ultraNumericEditor17 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label59 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtVestimenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label50 = new System.Windows.Forms.Label();
			this.txtEducacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label49 = new System.Windows.Forms.Label();
			this.txtSalud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label48 = new System.Windows.Forms.Label();
			this.txtAlimentacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label47 = new System.Windows.Forms.Label();
			this.txtVivienda = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label46 = new System.Windows.Forms.Label();
			this.ultraTabPageControl12 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid6 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGrid5 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsRolConcepto = new C1.Data.C1DataSet();
			this.cdvSucursal = new C1.Data.C1DataView();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtIdPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btNivel = new Infragistics.Win.Misc.UltraButton();
			this.btTitulo = new Infragistics.Win.Misc.UltraButton();
			this.btParentesco = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btCargasF = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.btAdicionales = new Infragistics.Win.Misc.UltraButton();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.lblPersonalDiscapacitado = new System.Windows.Forms.Label();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCargas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoSectorial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optSistema)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAfiliacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optSexo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReingreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolReingreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRegSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRegIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRelacionT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtNovedad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngreso)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbJefes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtJornal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSueldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalendario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDistrib)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsientoDistrib)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraOptionSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNumConadis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoDiscapacidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeDiscapacidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.ultraTabPageControl8.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbHastaV)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesdeV)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid4)).BeginInit();
			this.ultraTabPageControl11.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGasto)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor17)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtVestimenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEducacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlimentacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVivienda)).BeginInit();
			this.ultraTabPageControl12.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.cmbCedula);
			this.ultraTabPageControl1.Controls.Add(this.spnCargas);
			this.ultraTabPageControl1.Controls.Add(this.label8);
			this.ultraTabPageControl1.Controls.Add(this.picFoto);
			this.ultraTabPageControl1.Controls.Add(this.txtCodigo);
			this.ultraTabPageControl1.Controls.Add(this.label7);
			this.ultraTabPageControl1.Controls.Add(this.txtCodigoSectorial);
			this.ultraTabPageControl1.Controls.Add(this.label6);
			this.ultraTabPageControl1.Controls.Add(this.optSistema);
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Controls.Add(this.txtAfiliacion);
			this.ultraTabPageControl1.Controls.Add(this.label4);
			this.ultraTabPageControl1.Controls.Add(this.txtCedula);
			this.ultraTabPageControl1.Controls.Add(this.txtNombre);
			this.ultraTabPageControl1.Controls.Add(this.label1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(724, 335);
			// 
			// cmbCedula
			// 
			this.cmbCedula.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Tipo"));
			this.cmbCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCedula.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			this.cmbCedula.Enabled = false;
			valueListItem1.DataValue = 2;
			valueListItem1.DisplayText = "Cédula";
			valueListItem2.DataValue = 3;
			valueListItem2.DisplayText = "Pasaporte";
			this.cmbCedula.Items.Add(valueListItem1);
			this.cmbCedula.Items.Add(valueListItem2);
			this.cmbCedula.Location = new System.Drawing.Point(24, 43);
			this.cmbCedula.Name = "cmbCedula";
			this.cmbCedula.Size = new System.Drawing.Size(120, 22);
			this.cmbCedula.TabIndex = 19;
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.FillOnRequest = false;
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// spnCargas
			// 
			this.spnCargas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.CargasFamiliares"));
			this.spnCargas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCargas.Location = new System.Drawing.Point(176, 112);
			this.spnCargas.MinValue = 0;
			this.spnCargas.Name = "spnCargas";
			this.spnCargas.PromptChar = ' ';
			this.spnCargas.Size = new System.Drawing.Size(112, 22);
			this.spnCargas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnCargas.TabIndex = 18;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(312, 17);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(30, 16);
			this.label8.TabIndex = 17;
			this.label8.Text = "Foto:";
			this.label8.Visible = false;
			// 
			// picFoto
			// 
			this.picFoto.BorderShadowColor = System.Drawing.Color.Empty;
			this.picFoto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Inset;
			this.picFoto.Location = new System.Drawing.Point(368, 17);
			this.picFoto.Name = "picFoto";
			this.picFoto.Size = new System.Drawing.Size(160, 155);
			this.picFoto.TabIndex = 16;
			this.picFoto.Visible = false;
			this.picFoto.Click += new System.EventHandler(this.picFoto_Click);
			// 
			// txtCodigo
			// 
			appearance1.BackColor = System.Drawing.Color.Gainsboro;
			appearance1.BackColor2 = System.Drawing.Color.White;
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(176, 146);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 22);
			this.txtCodigo.TabIndex = 15;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(24, 146);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(117, 16);
			this.label7.TabIndex = 14;
			this.label7.Text = "Codigo del Empleado::";
			// 
			// txtCodigoSectorial
			// 
			appearance2.BackColor = System.Drawing.Color.Gainsboro;
			appearance2.BackColor2 = System.Drawing.Color.White;
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCodigoSectorial.Appearance = appearance2;
			this.txtCodigoSectorial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoSectorial.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.CodigoSectorial"));
			this.txtCodigoSectorial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoSectorial.Location = new System.Drawing.Point(416, 215);
			this.txtCodigoSectorial.Name = "txtCodigoSectorial";
			this.txtCodigoSectorial.Size = new System.Drawing.Size(112, 22);
			this.txtCodigoSectorial.TabIndex = 13;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(304, 215);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(89, 16);
			this.label6.TabIndex = 12;
			this.label6.Text = "Codigo Sectorial:";
			// 
			// optSistema
			// 
			this.optSistema.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.SistemaSalario"));
			this.optSistema.ItemAppearance = appearance3;
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Sin Sistema de Salario Neto";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Con Sistema de Salario Neto";
			this.optSistema.Items.Add(valueListItem3);
			this.optSistema.Items.Add(valueListItem4);
			this.optSistema.ItemSpacingVertical = 6;
			this.optSistema.Location = new System.Drawing.Point(80, 190);
			this.optSistema.Name = "optSistema";
			this.optSistema.Size = new System.Drawing.Size(208, 51);
			this.optSistema.TabIndex = 11;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(24, 112);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(67, 16);
			this.label5.TabIndex = 8;
			this.label5.Text = "Cargas Fam.";
			// 
			// txtAfiliacion
			// 
			appearance4.BackColor = System.Drawing.Color.Gainsboro;
			appearance4.BackColor2 = System.Drawing.Color.White;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtAfiliacion.Appearance = appearance4;
			this.txtAfiliacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtAfiliacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Afiliacion"));
			this.txtAfiliacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAfiliacion.Location = new System.Drawing.Point(176, 78);
			this.txtAfiliacion.Name = "txtAfiliacion";
			this.txtAfiliacion.Size = new System.Drawing.Size(112, 22);
			this.txtAfiliacion.TabIndex = 7;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(24, 78);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(49, 16);
			this.label4.TabIndex = 6;
			this.label4.Text = "Afiliación";
			// 
			// txtCedula
			// 
			appearance5.BackColor = System.Drawing.Color.Gainsboro;
			appearance5.BackColor2 = System.Drawing.Color.White;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCedula.Appearance = appearance5;
			this.txtCedula.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Cedula"));
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Location = new System.Drawing.Point(176, 43);
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(112, 22);
			this.txtCedula.TabIndex = 5;
			this.txtCedula.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtCedula_MouseDown);
			// 
			// txtNombre
			// 
			appearance6.BackColor = System.Drawing.Color.Gainsboro;
			appearance6.BackColor2 = System.Drawing.Color.White;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNombre.Appearance = appearance6;
			this.txtNombre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(80, 17);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(208, 20);
			this.txtNombre.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(24, 17);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Nombre:";
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.lblFiltro);
			this.ultraTabPageControl2.Controls.Add(this.cmbFiltro);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor11);
			this.ultraTabPageControl2.Controls.Add(this.label64);
			this.ultraTabPageControl2.Controls.Add(this.ultraNumericEditor8);
			this.ultraTabPageControl2.Controls.Add(this.label43);
			this.ultraTabPageControl2.Controls.Add(this.ultraNumericEditor7);
			this.ultraTabPageControl2.Controls.Add(this.label42);
			this.ultraTabPageControl2.Controls.Add(this.btProvincia);
			this.ultraTabPageControl2.Controls.Add(this.btCiudad);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor9);
			this.ultraTabPageControl2.Controls.Add(this.label40);
			this.ultraTabPageControl2.Controls.Add(this.optSexo);
			this.ultraTabPageControl2.Controls.Add(this.ultraCombo2);
			this.ultraTabPageControl2.Controls.Add(this.ultraCombo1);
			this.ultraTabPageControl2.Controls.Add(this.label22);
			this.ultraTabPageControl2.Controls.Add(this.label21);
			this.ultraTabPageControl2.Controls.Add(this.txtCelular);
			this.ultraTabPageControl2.Controls.Add(this.label20);
			this.ultraTabPageControl2.Controls.Add(this.txtTelefono);
			this.ultraTabPageControl2.Controls.Add(this.label19);
			this.ultraTabPageControl2.Controls.Add(this.txtDireccion);
			this.ultraTabPageControl2.Controls.Add(this.label18);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(724, 335);
			this.ultraTabPageControl2.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl2_Paint);
			// 
			// lblFiltro
			// 
			this.lblFiltro.AutoSize = true;
			this.lblFiltro.BackColor = System.Drawing.Color.Transparent;
			this.lblFiltro.Location = new System.Drawing.Point(280, 215);
			this.lblFiltro.Name = "lblFiltro";
			this.lblFiltro.Size = new System.Drawing.Size(68, 16);
			this.lblFiltro.TabIndex = 26;
			this.lblFiltro.Text = "Nuevo Filtro:";
			// 
			// cmbFiltro
			// 
			this.cmbFiltro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFiltro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFiltro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idCiudad"));
			this.cmbFiltro.DataMember = "ClienteCiudad";
			this.cmbFiltro.DataSource = this.cdsCiudad;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 166;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 81;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbFiltro.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbFiltro.DisplayMember = "Ciudad";
			this.cmbFiltro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFiltro.Location = new System.Drawing.Point(360, 207);
			this.cmbFiltro.Name = "cmbFiltro";
			this.cmbFiltro.Size = new System.Drawing.Size(120, 21);
			this.cmbFiltro.TabIndex = 25;
			this.cmbFiltro.ValueMember = "idCiudad";
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
			// ultraTextEditor11
			// 
			appearance7.BackColor = System.Drawing.Color.Gainsboro;
			appearance7.BackColor2 = System.Drawing.Color.White;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor11.Appearance = appearance7;
			this.ultraTextEditor11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.TipoSangre"));
			this.ultraTextEditor11.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor11.Location = new System.Drawing.Point(360, 164);
			this.ultraTextEditor11.Name = "ultraTextEditor11";
			this.ultraTextEditor11.Size = new System.Drawing.Size(120, 22);
			this.ultraTextEditor11.TabIndex = 24;
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.BackColor = System.Drawing.Color.Transparent;
			this.label64.Location = new System.Drawing.Point(280, 164);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(64, 16);
			this.label64.TabIndex = 23;
			this.label64.Text = "Tipo Sangre";
			// 
			// ultraNumericEditor8
			// 
			this.ultraNumericEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.NumeroDias"));
			this.ultraNumericEditor8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor8.Location = new System.Drawing.Point(112, 233);
			this.ultraNumericEditor8.MinValue = 0;
			this.ultraNumericEditor8.Name = "ultraNumericEditor8";
			this.ultraNumericEditor8.PromptChar = ' ';
			this.ultraNumericEditor8.Size = new System.Drawing.Size(112, 22);
			this.ultraNumericEditor8.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor8.TabIndex = 22;
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.BackColor = System.Drawing.Color.Transparent;
			this.label43.Location = new System.Drawing.Point(24, 233);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(67, 16);
			this.label43.TabIndex = 21;
			this.label43.Text = "Dias Trabajo";
			// 
			// ultraNumericEditor7
			// 
			this.ultraNumericEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Educacion"));
			this.ultraNumericEditor7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor7.Location = new System.Drawing.Point(112, 198);
			this.ultraNumericEditor7.MinValue = 0;
			this.ultraNumericEditor7.Name = "ultraNumericEditor7";
			this.ultraNumericEditor7.PromptChar = ' ';
			this.ultraNumericEditor7.Size = new System.Drawing.Size(112, 22);
			this.ultraNumericEditor7.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor7.TabIndex = 20;
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(24, 198);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(55, 16);
			this.label42.TabIndex = 19;
			this.label42.Text = "Educacion";
			// 
			// btProvincia
			// 
			this.btProvincia.Location = new System.Drawing.Point(488, 95);
			this.btProvincia.Name = "btProvincia";
			this.btProvincia.Size = new System.Drawing.Size(24, 25);
			this.btProvincia.TabIndex = 17;
			this.btProvincia.Text = "...";
			this.btProvincia.Click += new System.EventHandler(this.btProvincia_Click);
			// 
			// btCiudad
			// 
			this.btCiudad.Location = new System.Drawing.Point(240, 95);
			this.btCiudad.Name = "btCiudad";
			this.btCiudad.Size = new System.Drawing.Size(24, 25);
			this.btCiudad.TabIndex = 16;
			this.btCiudad.Text = "...";
			this.btCiudad.Click += new System.EventHandler(this.btCiudad_Click);
			// 
			// ultraTextEditor9
			// 
			appearance8.BackColor = System.Drawing.Color.Gainsboro;
			appearance8.BackColor2 = System.Drawing.Color.White;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor9.Appearance = appearance8;
			this.ultraTextEditor9.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Numero"));
			this.ultraTextEditor9.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor9.Location = new System.Drawing.Point(360, 129);
			this.ultraTextEditor9.Name = "ultraTextEditor9";
			this.ultraTextEditor9.Size = new System.Drawing.Size(120, 22);
			this.ultraTextEditor9.TabIndex = 15;
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Location = new System.Drawing.Point(280, 129);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(44, 16);
			this.label40.TabIndex = 14;
			this.label40.Text = "Numero";
			// 
			// optSexo
			// 
			this.optSexo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Sexo"));
			this.optSexo.ItemAppearance = appearance9;
			this.optSexo.ItemOrigin = new System.Drawing.Point(10, 2);
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Hombre";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Mujer";
			this.optSexo.Items.Add(valueListItem5);
			this.optSexo.Items.Add(valueListItem6);
			this.optSexo.ItemSpacingHorizontal = 10;
			this.optSexo.ItemSpacingVertical = 6;
			this.optSexo.Location = new System.Drawing.Point(112, 129);
			this.optSexo.Name = "optSexo";
			this.optSexo.Size = new System.Drawing.Size(96, 52);
			this.optSexo.TabIndex = 13;
			// 
			// ultraCombo2
			// 
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idProvincia"));
			this.ultraCombo2.DataMember = "Provincia";
			this.ultraCombo2.DataSource = this.cdsProvincia;
			ultraGridBand2.AddButtonCaption = "ClienteCiudad";
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 79;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 156;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ultraCombo2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraCombo2.DisplayMember = "provincia";
			this.ultraCombo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(360, 95);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(120, 21);
			this.ultraCombo2.TabIndex = 12;
			this.ultraCombo2.ValueMember = "idProvincia";
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
			// ultraCombo1
			// 
			this.ultraCombo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idCiudad"));
			this.ultraCombo1.DataMember = "ClienteCiudad";
			this.ultraCombo1.DataSource = this.cdsCiudad;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 166;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 81;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ultraCombo1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraCombo1.DisplayMember = "Ciudad";
			this.ultraCombo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo1.Location = new System.Drawing.Point(112, 95);
			this.ultraCombo1.Name = "ultraCombo1";
			this.ultraCombo1.Size = new System.Drawing.Size(120, 21);
			this.ultraCombo1.TabIndex = 11;
			this.ultraCombo1.ValueMember = "idCiudad";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(280, 95);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(53, 16);
			this.label22.TabIndex = 10;
			this.label22.Text = "Provincia:";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(24, 95);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(42, 16);
			this.label21.TabIndex = 8;
			this.label21.Text = "Ciudad:";
			// 
			// txtCelular
			// 
			appearance10.BackColor = System.Drawing.Color.Gainsboro;
			appearance10.BackColor2 = System.Drawing.Color.White;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCelular.Appearance = appearance10;
			this.txtCelular.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Celular"));
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Location = new System.Drawing.Point(360, 60);
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(120, 22);
			this.txtCelular.TabIndex = 7;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(280, 60);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(42, 16);
			this.label20.TabIndex = 6;
			this.label20.Text = "Celular:";
			// 
			// txtTelefono
			// 
			appearance11.BackColor = System.Drawing.Color.Gainsboro;
			appearance11.BackColor2 = System.Drawing.Color.White;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtTelefono.Appearance = appearance11;
			this.txtTelefono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Telefono"));
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Location = new System.Drawing.Point(112, 60);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(120, 22);
			this.txtTelefono.TabIndex = 5;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(24, 60);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(52, 16);
			this.label19.TabIndex = 4;
			this.label19.Text = "Teléfono:";
			// 
			// txtDireccion
			// 
			appearance12.BackColor = System.Drawing.Color.Gainsboro;
			appearance12.BackColor2 = System.Drawing.Color.White;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtDireccion.Appearance = appearance12;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Direccion"));
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Location = new System.Drawing.Point(112, 26);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(368, 22);
			this.txtDireccion.TabIndex = 3;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(24, 26);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(50, 16);
			this.label18.TabIndex = 2;
			this.label18.Text = "Dirección";
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.label12);
			this.ultraTabPageControl3.Controls.Add(this.ultraCalendarCombo2);
			this.ultraTabPageControl3.Controls.Add(this.grdReingreso);
			this.ultraTabPageControl3.Controls.Add(this.btnReingreso);
			this.ultraTabPageControl3.Controls.Add(this.label63);
			this.ultraTabPageControl3.Controls.Add(this.ultraButton2);
			this.ultraTabPageControl3.Controls.Add(this.ultraCombo5);
			this.ultraTabPageControl3.Controls.Add(this.label62);
			this.ultraTabPageControl3.Controls.Add(this.txtRegSalida);
			this.ultraTabPageControl3.Controls.Add(this.label61);
			this.ultraTabPageControl3.Controls.Add(this.txtRegIngreso);
			this.ultraTabPageControl3.Controls.Add(this.ultraCombo4);
			this.ultraTabPageControl3.Controls.Add(this.label45);
			this.ultraTabPageControl3.Controls.Add(this.label44);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor10);
			this.ultraTabPageControl3.Controls.Add(this.label41);
			this.ultraTabPageControl3.Controls.Add(this.ultraCalendarCombo1);
			this.ultraTabPageControl3.Controls.Add(this.label15);
			this.ultraTabPageControl3.Controls.Add(this.dtNovedad);
			this.ultraTabPageControl3.Controls.Add(this.label11);
			this.ultraTabPageControl3.Controls.Add(this.dtSalida);
			this.ultraTabPageControl3.Controls.Add(this.label10);
			this.ultraTabPageControl3.Controls.Add(this.dtNacimiento);
			this.ultraTabPageControl3.Controls.Add(this.label9);
			this.ultraTabPageControl3.Controls.Add(this.dtIngreso);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(724, 335);
			this.ultraTabPageControl3.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl3_Paint);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(32, 200);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(106, 16);
			this.label12.TabIndex = 25;
			this.label12.Text = "Fecha de Reingreso:";
			// 
			// ultraCalendarCombo2
			// 
			this.ultraCalendarCombo2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCalendarCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.FechaReIngreso"));
			dateButton1.Caption = "Today";
			this.ultraCalendarCombo2.DateButtons.Add(dateButton1);
			this.ultraCalendarCombo2.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo2.Location = new System.Drawing.Point(160, 200);
			this.ultraCalendarCombo2.Name = "ultraCalendarCombo2";
			this.ultraCalendarCombo2.NonAutoSizeHeight = 23;
			this.ultraCalendarCombo2.Size = new System.Drawing.Size(128, 21);
			this.ultraCalendarCombo2.TabIndex = 24;
			this.ultraCalendarCombo2.Value = new System.DateTime(2005, 5, 23, 0, 0, 0, 0);
			// 
			// grdReingreso
			// 
			this.grdReingreso.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdReingreso.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdReingreso.DataMember = "RolReingreso";
			this.grdReingreso.DataSource = this.cdsRolReingreso;
			appearance13.BackColor = System.Drawing.Color.White;
			this.grdReingreso.DisplayLayout.Appearance = appearance13;
			this.grdReingreso.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 146;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 140;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 99;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.grdReingreso.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.grdReingreso.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdReingreso.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdReingreso.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.grdReingreso.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 10F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdReingreso.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdReingreso.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdReingreso.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.grdReingreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdReingreso.Location = new System.Drawing.Point(160, 240);
			this.grdReingreso.Name = "grdReingreso";
			this.grdReingreso.Size = new System.Drawing.Size(423, 87);
			this.grdReingreso.TabIndex = 23;
			this.grdReingreso.Visible = false;
			// 
			// cdsRolReingreso
			// 
			this.cdsRolReingreso.BindingContextCtrl = this;
			this.cdsRolReingreso.DataLibrary = "LibPersonal";
			this.cdsRolReingreso.DataLibraryUrl = "";
			this.cdsRolReingreso.DataSetDef = "dsRolTablas";
			this.cdsRolReingreso.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolReingreso.Name = "cdsRolReingreso";
			this.cdsRolReingreso.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolReingreso.SchemaDef = null;
			this.cdsRolReingreso.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolReingreso_BeforeFill);
			// 
			// btnReingreso
			// 
			this.btnReingreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnReingreso.Location = new System.Drawing.Point(48, 272);
			this.btnReingreso.Name = "btnReingreso";
			this.btnReingreso.Size = new System.Drawing.Size(88, 25);
			this.btnReingreso.TabIndex = 22;
			this.btnReingreso.Text = "Reingreso";
			this.btnReingreso.Visible = false;
			this.btnReingreso.Click += new System.EventHandler(this.btnReingreso_Click);
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.BackColor = System.Drawing.Color.Transparent;
			this.label63.Location = new System.Drawing.Point(336, 164);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(72, 16);
			this.label63.TabIndex = 21;
			this.label63.Text = "Tipo Contrato";
			this.label63.Click += new System.EventHandler(this.label63_Click);
			// 
			// ultraButton2
			// 
			this.ultraButton2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton2.Location = new System.Drawing.Point(584, 164);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(24, 24);
			this.ultraButton2.TabIndex = 20;
			this.ultraButton2.Text = "...";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// ultraCombo5
			// 
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance18.BackColor2 = System.Drawing.Color.White;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraCombo5.Appearance = appearance18;
			this.ultraCombo5.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo5.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idTipoContrato"));
			this.ultraCombo5.DataMember = "RolTipoContrato";
			this.ultraCombo5.DataSource = this.cdsTipoContrato;
			ultraGridBand5.AddButtonCaption = "RolSucursal";
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.ultraCombo5.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ultraCombo5.DisplayMember = "TipoContrato";
			this.ultraCombo5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo5.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo5.Location = new System.Drawing.Point(432, 164);
			this.ultraCombo5.Name = "ultraCombo5";
			this.ultraCombo5.Size = new System.Drawing.Size(144, 21);
			this.ultraCombo5.TabIndex = 19;
			this.ultraCombo5.ValueMember = "idTipoContrato";
			// 
			// cdsTipoContrato
			// 
			this.cdsTipoContrato.BindingContextCtrl = this;
			this.cdsTipoContrato.DataLibrary = "LibPersonal";
			this.cdsTipoContrato.DataLibraryUrl = "";
			this.cdsTipoContrato.DataSetDef = "dsRolTablas";
			this.cdsTipoContrato.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoContrato.Name = "cdsTipoContrato";
			this.cdsTipoContrato.SchemaClassName = "LibPersonal.DataClass";
			this.cdsTipoContrato.SchemaDef = null;
			this.cdsTipoContrato.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoContrato_BeforeFill);
			// 
			// label62
			// 
			this.label62.AutoSize = true;
			this.label62.BackColor = System.Drawing.Color.Transparent;
			this.label62.Location = new System.Drawing.Point(336, 95);
			this.label62.Name = "label62";
			this.label62.Size = new System.Drawing.Size(75, 16);
			this.label62.TabIndex = 18;
			this.label62.Text = "Reg. N. Salida";
			// 
			// txtRegSalida
			// 
			appearance19.BackColor = System.Drawing.Color.Gainsboro;
			appearance19.BackColor2 = System.Drawing.Color.White;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtRegSalida.Appearance = appearance19;
			this.txtRegSalida.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRegSalida.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.RegSalida"));
			this.txtRegSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRegSalida.Location = new System.Drawing.Point(432, 95);
			this.txtRegSalida.Name = "txtRegSalida";
			this.txtRegSalida.Size = new System.Drawing.Size(144, 22);
			this.txtRegSalida.TabIndex = 17;
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(336, 26);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(84, 16);
			this.label61.TabIndex = 16;
			this.label61.Text = "Reg. N. Entrada";
			// 
			// txtRegIngreso
			// 
			appearance20.BackColor = System.Drawing.Color.Gainsboro;
			appearance20.BackColor2 = System.Drawing.Color.White;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtRegIngreso.Appearance = appearance20;
			this.txtRegIngreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRegIngreso.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.RegEntrada"));
			this.txtRegIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRegIngreso.Location = new System.Drawing.Point(432, 26);
			this.txtRegIngreso.Name = "txtRegIngreso";
			this.txtRegIngreso.Size = new System.Drawing.Size(144, 22);
			this.txtRegIngreso.TabIndex = 15;
			// 
			// ultraCombo4
			// 
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance21.BackColor2 = System.Drawing.Color.White;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraCombo4.Appearance = appearance21;
			this.ultraCombo4.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo4.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idRelacionT"));
			this.ultraCombo4.DataMember = "RolRelacionTrabajo";
			this.ultraCombo4.DataSource = this.cdsRelacionT;
			ultraGridBand6.AddButtonCaption = "RolSucursal";
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 66;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 223;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.ultraCombo4.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraCombo4.DisplayMember = "Relacion";
			this.ultraCombo4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo4.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo4.Location = new System.Drawing.Point(432, 60);
			this.ultraCombo4.Name = "ultraCombo4";
			this.ultraCombo4.Size = new System.Drawing.Size(144, 21);
			this.ultraCombo4.TabIndex = 14;
			this.ultraCombo4.ValueMember = "idRelacionTrabajo";
			// 
			// cdsRelacionT
			// 
			this.cdsRelacionT.BindingContextCtrl = this;
			this.cdsRelacionT.DataLibrary = "LibPersonal";
			this.cdsRelacionT.DataLibraryUrl = "";
			this.cdsRelacionT.DataSetDef = "dsRolTablas";
			this.cdsRelacionT.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRelacionT.Name = "cdsRelacionT";
			this.cdsRelacionT.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRelacionT.SchemaDef = null;
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Location = new System.Drawing.Point(336, 60);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(57, 16);
			this.label45.TabIndex = 13;
			this.label45.Text = "Empleador";
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Location = new System.Drawing.Point(336, 129);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(44, 16);
			this.label44.TabIndex = 11;
			this.label44.Text = "Jornada";
			// 
			// ultraTextEditor10
			// 
			appearance22.BackColor = System.Drawing.Color.Gainsboro;
			appearance22.BackColor2 = System.Drawing.Color.White;
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor10.Appearance = appearance22;
			this.ultraTextEditor10.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Jornada"));
			this.ultraTextEditor10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor10.Location = new System.Drawing.Point(432, 129);
			this.ultraTextEditor10.Name = "ultraTextEditor10";
			this.ultraTextEditor10.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor10.TabIndex = 10;
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Location = new System.Drawing.Point(32, 164);
			this.label41.Name = "label41";
			this.label41.TabIndex = 9;
			this.label41.Text = "Fecha de Contrato:";
			// 
			// ultraCalendarCombo1
			// 
			this.ultraCalendarCombo1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCalendarCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.FechaContrato"));
			dateButton2.Caption = "Today";
			this.ultraCalendarCombo1.DateButtons.Add(dateButton2);
			this.ultraCalendarCombo1.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo1.Location = new System.Drawing.Point(160, 164);
			this.ultraCalendarCombo1.Name = "ultraCalendarCombo1";
			this.ultraCalendarCombo1.NonAutoSizeHeight = 23;
			this.ultraCalendarCombo1.Size = new System.Drawing.Size(128, 21);
			this.ultraCalendarCombo1.TabIndex = 8;
			this.ultraCalendarCombo1.Value = new System.DateTime(2005, 5, 23, 0, 0, 0, 0);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(32, 129);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(101, 16);
			this.label15.TabIndex = 7;
			this.label15.Text = "Fecha de Novedad:";
			// 
			// dtNovedad
			// 
			this.dtNovedad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtNovedad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.FechaNovedad"));
			dateButton3.Caption = "Today";
			this.dtNovedad.DateButtons.Add(dateButton3);
			this.dtNovedad.Format = "dd/MMM/yyyy";
			this.dtNovedad.Location = new System.Drawing.Point(160, 129);
			this.dtNovedad.Name = "dtNovedad";
			this.dtNovedad.NonAutoSizeHeight = 23;
			this.dtNovedad.Size = new System.Drawing.Size(128, 21);
			this.dtNovedad.TabIndex = 6;
			this.dtNovedad.Value = new System.DateTime(2005, 5, 23, 0, 0, 0, 0);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(32, 95);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(86, 16);
			this.label11.TabIndex = 5;
			this.label11.Text = "Fecha de Salida:";
			// 
			// dtSalida
			// 
			this.dtSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtSalida.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.FechaSalida"));
			dateButton4.Caption = "Today";
			this.dtSalida.DateButtons.Add(dateButton4);
			this.dtSalida.Format = "dd/MMM/yyyy";
			this.dtSalida.Location = new System.Drawing.Point(160, 95);
			this.dtSalida.Name = "dtSalida";
			this.dtSalida.NonAutoSizeHeight = 23;
			this.dtSalida.Size = new System.Drawing.Size(128, 21);
			this.dtSalida.TabIndex = 4;
			this.dtSalida.Value = new System.DateTime(2005, 5, 23, 0, 0, 0, 0);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(32, 60);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(112, 16);
			this.label10.TabIndex = 3;
			this.label10.Text = "Fecha de Nacimiento:";
			// 
			// dtNacimiento
			// 
			this.dtNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtNacimiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.FechaNacimiento"));
			dateButton5.Caption = "Today";
			this.dtNacimiento.DateButtons.Add(dateButton5);
			this.dtNacimiento.Format = "dd/MMM/yyyy";
			this.dtNacimiento.Location = new System.Drawing.Point(160, 60);
			this.dtNacimiento.Name = "dtNacimiento";
			this.dtNacimiento.NonAutoSizeHeight = 23;
			this.dtNacimiento.Size = new System.Drawing.Size(128, 21);
			this.dtNacimiento.TabIndex = 2;
			this.dtNacimiento.Value = new System.DateTime(2005, 5, 23, 0, 0, 0, 0);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(32, 26);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(95, 16);
			this.label9.TabIndex = 1;
			this.label9.Text = "Fecha de Ingreso:";
			// 
			// dtIngreso
			// 
			this.dtIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtIngreso.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.FechaIngreso"));
			dateButton6.Caption = "Today";
			this.dtIngreso.DateButtons.Add(dateButton6);
			this.dtIngreso.Format = "dd/MMM/yyyy";
			this.dtIngreso.Location = new System.Drawing.Point(160, 26);
			this.dtIngreso.Name = "dtIngreso";
			this.dtIngreso.NonAutoSizeHeight = 23;
			this.dtIngreso.Size = new System.Drawing.Size(128, 21);
			this.dtIngreso.TabIndex = 0;
			this.dtIngreso.Value = new System.DateTime(2005, 5, 23, 0, 0, 0, 0);
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.cmbJefes);
			this.ultraTabPageControl4.Controls.Add(this.label66);
			this.ultraTabPageControl4.Controls.Add(this.chkSupervisor);
			this.ultraTabPageControl4.Controls.Add(this.cmbSupervisor);
			this.ultraTabPageControl4.Controls.Add(this.label65);
			this.ultraTabPageControl4.Controls.Add(this.label3);
			this.ultraTabPageControl4.Controls.Add(this.ultraNumericEditor6);
			this.ultraTabPageControl4.Controls.Add(this.cmbCuenta);
			this.ultraTabPageControl4.Controls.Add(this.label34);
			this.ultraTabPageControl4.Controls.Add(this.cmbCentroCosto);
			this.ultraTabPageControl4.Controls.Add(this.label33);
			this.ultraTabPageControl4.Controls.Add(this.ultraNumericEditor1);
			this.ultraTabPageControl4.Controls.Add(this.label27);
			this.ultraTabPageControl4.Controls.Add(this.chkCobrador);
			this.ultraTabPageControl4.Controls.Add(this.chkVendedor);
			this.ultraTabPageControl4.Controls.Add(this.btCargo);
			this.ultraTabPageControl4.Controls.Add(this.btDepartamento);
			this.ultraTabPageControl4.Controls.Add(this.txtJornal);
			this.ultraTabPageControl4.Controls.Add(this.label17);
			this.ultraTabPageControl4.Controls.Add(this.txtSueldo);
			this.ultraTabPageControl4.Controls.Add(this.label16);
			this.ultraTabPageControl4.Controls.Add(this.chkFondoReserva);
			this.ultraTabPageControl4.Controls.Add(this.chkRenta);
			this.ultraTabPageControl4.Controls.Add(this.chkIess);
			this.ultraTabPageControl4.Controls.Add(this.cmbCargo);
			this.ultraTabPageControl4.Controls.Add(this.label14);
			this.ultraTabPageControl4.Controls.Add(this.cmbDepartamento);
			this.ultraTabPageControl4.Controls.Add(this.label13);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(724, 335);
			// 
			// cmbJefes
			// 
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance23.BackColor2 = System.Drawing.Color.White;
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbJefes.Appearance = appearance23;
			this.cmbJefes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbJefes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbJefes.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idJefe"));
			this.cmbJefes.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbJefes.DisplayMember = "Descripcion";
			this.cmbJefes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbJefes.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbJefes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbJefes.Location = new System.Drawing.Point(136, 216);
			this.cmbJefes.Name = "cmbJefes";
			this.cmbJefes.Size = new System.Drawing.Size(208, 21);
			this.cmbJefes.TabIndex = 36;
			this.cmbJefes.ValueMember = "";
			this.cmbJefes.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbJefes_MouseUp);
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.BackColor = System.Drawing.Color.Transparent;
			this.label66.Location = new System.Drawing.Point(16, 218);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(68, 16);
			this.label66.TabIndex = 35;
			this.label66.Text = "Depende de:";
			// 
			// chkSupervisor
			// 
			this.chkSupervisor.BackColor = System.Drawing.Color.Transparent;
			this.chkSupervisor.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.bSupervisor"));
			this.chkSupervisor.Location = new System.Drawing.Point(416, 240);
			this.chkSupervisor.Name = "chkSupervisor";
			this.chkSupervisor.Size = new System.Drawing.Size(136, 22);
			this.chkSupervisor.TabIndex = 34;
			this.chkSupervisor.Text = "Supervisor";
			this.chkSupervisor.CheckedChanged += new System.EventHandler(this.chkSupervisor_CheckedChanged);
			// 
			// cmbSupervisor
			// 
			this.cmbSupervisor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSupervisor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSupervisor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idSupervisor"));
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbSupervisor.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSupervisor.DisplayMember = "Nombre";
			this.cmbSupervisor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSupervisor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSupervisor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSupervisor.Location = new System.Drawing.Point(136, 184);
			this.cmbSupervisor.Name = "cmbSupervisor";
			this.cmbSupervisor.Size = new System.Drawing.Size(208, 21);
			this.cmbSupervisor.TabIndex = 32;
			this.cmbSupervisor.ValueMember = "idPersonal";
			// 
			// label65
			// 
			this.label65.AutoSize = true;
			this.label65.BackColor = System.Drawing.Color.Transparent;
			this.label65.Location = new System.Drawing.Point(16, 186);
			this.label65.Name = "label65";
			this.label65.Size = new System.Drawing.Size(85, 16);
			this.label65.TabIndex = 31;
			this.label65.Text = "Supervisado por";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(232, 147);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 30;
			this.label3.Text = "Convenio";
			// 
			// ultraNumericEditor6
			// 
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance24.BackColor2 = System.Drawing.Color.White;
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraNumericEditor6.Appearance = appearance24;
			this.ultraNumericEditor6.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraNumericEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Convenio"));
			this.ultraNumericEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor6.FormatString = "#,##0.00";
			this.ultraNumericEditor6.Location = new System.Drawing.Point(288, 144);
			this.ultraNumericEditor6.Name = "ultraNumericEditor6";
			this.ultraNumericEditor6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor6.PromptChar = ' ';
			this.ultraNumericEditor6.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor6.TabIndex = 29;
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idCuenta"));
			this.cmbCuenta.DataSource = this.cdvCuenta;
			ultraGridBand8.AddButtonCaption = "RolTipo";
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Width = 96;
			ultraGridColumn27.Header.VisiblePosition = 12;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 8;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Width = 157;
			ultraGridColumn29.Header.VisiblePosition = 9;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 8;
			ultraGridColumn30.Header.VisiblePosition = 4;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 8;
			ultraGridColumn31.Header.VisiblePosition = 6;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 8;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 17;
			ultraGridColumn33.Header.VisiblePosition = 3;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 17;
			ultraGridColumn34.Header.VisiblePosition = 7;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 17;
			ultraGridColumn35.Header.VisiblePosition = 10;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 8;
			ultraGridColumn36.Header.VisiblePosition = 11;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 17;
			ultraGridColumn37.Header.VisiblePosition = 13;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 17;
			ultraGridColumn38.Header.VisiblePosition = 14;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 17;
			ultraGridColumn39.Header.VisiblePosition = 15;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 17;
			ultraGridColumn40.Header.VisiblePosition = 16;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 109;
			ultraGridColumn41.Header.VisiblePosition = 17;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 153;
			ultraGridColumn42.Header.VisiblePosition = 18;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 146;
			ultraGridColumn43.Header.VisiblePosition = 5;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 8;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 232;
			ultraGridColumn45.Header.VisiblePosition = 19;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn45});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCuenta.DisplayMember = "Codigo";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(136, 112);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(208, 21);
			this.cmbCuenta.TabIndex = 28;
			this.cmbCuenta.ValueMember = "idCuenta";
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(16, 114);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(39, 16);
			this.label34.TabIndex = 27;
			this.label34.Text = "Cuenta";
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idProyecto"));
			this.cmbCentroCosto.DataSource = this.cdvProyecto;
			ultraGridBand9.AddButtonCaption = "RolTipo";
			ultraGridColumn46.Header.VisiblePosition = 2;
			ultraGridColumn46.Width = 85;
			ultraGridColumn47.Header.VisiblePosition = 6;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 14;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Width = 311;
			ultraGridColumn49.Header.VisiblePosition = 3;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 24;
			ultraGridColumn50.Header.VisiblePosition = 7;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 24;
			ultraGridColumn51.Header.VisiblePosition = 4;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 24;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 94;
			ultraGridColumn53.Header.VisiblePosition = 5;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 22;
			ultraGridColumn54.Header.VisiblePosition = 8;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 17;
			ultraGridColumn55.Header.VisiblePosition = 9;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 10;
			ultraGridColumn57.Header.VisiblePosition = 11;
			ultraGridColumn58.Header.VisiblePosition = 12;
			ultraGridBand9.Columns.AddRange(new object[] {
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
																										 ultraGridColumn58});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(136, 79);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(208, 21);
			this.cmbCentroCosto.TabIndex = 26;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
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
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(16, 81);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(69, 16);
			this.label33.TabIndex = 24;
			this.label33.Text = "Centro Costo";
			// 
			// ultraNumericEditor1
			// 
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance25.BackColor2 = System.Drawing.Color.White;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraNumericEditor1.Appearance = appearance25;
			this.ultraNumericEditor1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Quincena"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(480, 208);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor1.TabIndex = 23;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(416, 211);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(51, 16);
			this.label27.TabIndex = 22;
			this.label27.Text = "Quincena";
			// 
			// chkCobrador
			// 
			this.chkCobrador.BackColor = System.Drawing.Color.Transparent;
			this.chkCobrador.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Cobrador"));
			this.chkCobrador.Location = new System.Drawing.Point(416, 144);
			this.chkCobrador.Name = "chkCobrador";
			this.chkCobrador.Size = new System.Drawing.Size(136, 21);
			this.chkCobrador.TabIndex = 17;
			this.chkCobrador.Text = "Cobrador";
			// 
			// chkVendedor
			// 
			this.chkVendedor.BackColor = System.Drawing.Color.Transparent;
			this.chkVendedor.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Vendedor"));
			this.chkVendedor.Location = new System.Drawing.Point(416, 112);
			this.chkVendedor.Name = "chkVendedor";
			this.chkVendedor.Size = new System.Drawing.Size(136, 22);
			this.chkVendedor.TabIndex = 16;
			this.chkVendedor.Text = "Vendedor";
			// 
			// btCargo
			// 
			this.btCargo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCargo.Location = new System.Drawing.Point(352, 48);
			this.btCargo.Name = "btCargo";
			this.btCargo.Size = new System.Drawing.Size(24, 25);
			this.btCargo.TabIndex = 15;
			this.btCargo.Text = "...";
			this.btCargo.Click += new System.EventHandler(this.btCargo_Click);
			// 
			// btDepartamento
			// 
			this.btDepartamento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btDepartamento.Location = new System.Drawing.Point(352, 16);
			this.btDepartamento.Name = "btDepartamento";
			this.btDepartamento.Size = new System.Drawing.Size(24, 25);
			this.btDepartamento.TabIndex = 14;
			this.btDepartamento.Text = "...";
			this.btDepartamento.Click += new System.EventHandler(this.btDepartamento_Click);
			// 
			// txtJornal
			// 
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance26.BackColor2 = System.Drawing.Color.White;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtJornal.Appearance = appearance26;
			this.txtJornal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtJornal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Jornal"));
			this.txtJornal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtJornal.FormatString = "#,##0.00";
			this.txtJornal.Location = new System.Drawing.Point(480, 176);
			this.txtJornal.Name = "txtJornal";
			this.txtJornal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtJornal.PromptChar = ' ';
			this.txtJornal.Size = new System.Drawing.Size(100, 22);
			this.txtJornal.TabIndex = 12;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(416, 179);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(38, 16);
			this.label17.TabIndex = 11;
			this.label17.Text = "Jornal:";
			// 
			// txtSueldo
			// 
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance27.BackColor2 = System.Drawing.Color.White;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtSueldo.Appearance = appearance27;
			this.txtSueldo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtSueldo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Sueldo"));
			this.txtSueldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSueldo.FormatString = "#,##0.00";
			this.txtSueldo.Location = new System.Drawing.Point(136, 144);
			this.txtSueldo.Name = "txtSueldo";
			this.txtSueldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSueldo.PromptChar = ' ';
			this.txtSueldo.Size = new System.Drawing.Size(88, 22);
			this.txtSueldo.TabIndex = 10;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(16, 147);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(42, 16);
			this.label16.TabIndex = 9;
			this.label16.Text = "Sueldo:";
			// 
			// chkFondoReserva
			// 
			this.chkFondoReserva.BackColor = System.Drawing.Color.Transparent;
			this.chkFondoReserva.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Reserva"));
			this.chkFondoReserva.Location = new System.Drawing.Point(416, 49);
			this.chkFondoReserva.Name = "chkFondoReserva";
			this.chkFondoReserva.Size = new System.Drawing.Size(240, 22);
			this.chkFondoReserva.TabIndex = 8;
			this.chkFondoReserva.Text = "Fondo de Reserva Cobrado Mensualmente";
			// 
			// chkRenta
			// 
			this.chkRenta.BackColor = System.Drawing.Color.Transparent;
			this.chkRenta.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Renta"));
			this.chkRenta.Location = new System.Drawing.Point(416, 80);
			this.chkRenta.Name = "chkRenta";
			this.chkRenta.Size = new System.Drawing.Size(136, 21);
			this.chkRenta.TabIndex = 7;
			this.chkRenta.Text = "Impuesto a la Renta";
			// 
			// chkIess
			// 
			this.chkIess.BackColor = System.Drawing.Color.Transparent;
			this.chkIess.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Iess"));
			this.chkIess.Location = new System.Drawing.Point(416, 18);
			this.chkIess.Name = "chkIess";
			this.chkIess.Size = new System.Drawing.Size(120, 21);
			this.chkIess.TabIndex = 6;
			this.chkIess.Text = "Aporte al IESS";
			// 
			// cmbCargo
			// 
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance28.BackColor2 = System.Drawing.Color.White;
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCargo.Appearance = appearance28;
			this.cmbCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idCargo"));
			this.cmbCargo.DataMember = "RolCargo";
			this.cmbCargo.DataSource = this.cdsSucursal;
			ultraGridBand10.AddButtonCaption = "RolSucursal";
			ultraGridColumn59.Header.VisiblePosition = 3;
			ultraGridColumn59.Width = 125;
			ultraGridColumn60.Header.VisiblePosition = 0;
			ultraGridColumn60.Width = 215;
			ultraGridColumn61.Header.VisiblePosition = 1;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn62.Hidden = true;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn59,
																											ultraGridColumn60,
																											ultraGridColumn61,
																											ultraGridColumn62});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbCargo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(136, 50);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(208, 21);
			this.cmbCargo.TabIndex = 5;
			this.cmbCargo.ValueMember = "idCargo";
			// 
			// cdsSucursal
			// 
			this.cdsSucursal.BindingContextCtrl = this;
			this.cdsSucursal.DataLibrary = "LibPersonal";
			this.cdsSucursal.DataLibraryUrl = "";
			this.cdsSucursal.DataSetDef = "dsRolTablas";
			this.cdsSucursal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSucursal.Name = "cdsSucursal";
			this.cdsSucursal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsSucursal.SchemaDef = null;
			this.cdsSucursal.BeforeFill += new C1.Data.FillEventHandler(this.cdsSucursal_BeforeFill);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(16, 52);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(37, 16);
			this.label14.TabIndex = 4;
			this.label14.Text = "Cargo:";
			// 
			// cmbDepartamento
			// 
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance29.BackColor2 = System.Drawing.Color.White;
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbDepartamento.Appearance = appearance29;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idDepartamento"));
			this.cmbDepartamento.DataMember = "RolDepartamento";
			this.cmbDepartamento.DataSource = this.cdsSucursal;
			ultraGridBand11.AddButtonCaption = "RolSucursal";
			ultraGridColumn63.Header.VisiblePosition = 2;
			ultraGridColumn63.Width = 111;
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Width = 271;
			ultraGridColumn65.Header.VisiblePosition = 1;
			ultraGridColumn65.Hidden = true;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn63,
																											ultraGridColumn64,
																											ultraGridColumn65});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(136, 18);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(208, 21);
			this.cmbDepartamento.TabIndex = 3;
			this.cmbDepartamento.ValueMember = "idDepartamento";
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(16, 20);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(80, 16);
			this.label13.TabIndex = 2;
			this.label13.Text = "Departamento:";
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.cmbBanco);
			this.ultraTabPageControl5.Controls.Add(this.btHoraExtra);
			this.ultraTabPageControl5.Controls.Add(this.btHorario1);
			this.ultraTabPageControl5.Controls.Add(this.btDistribucion);
			this.ultraTabPageControl5.Controls.Add(this.btHorario);
			this.ultraTabPageControl5.Controls.Add(this.cmbCalendario);
			this.ultraTabPageControl5.Controls.Add(this.label39);
			this.ultraTabPageControl5.Controls.Add(this.cmbDistrib);
			this.ultraTabPageControl5.Controls.Add(this.ultraTextEditor4);
			this.ultraTabPageControl5.Controls.Add(this.label26);
			this.ultraTabPageControl5.Controls.Add(this.ultraTextEditor3);
			this.ultraTabPageControl5.Controls.Add(this.label25);
			this.ultraTabPageControl5.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl5.Controls.Add(this.label24);
			this.ultraTabPageControl5.Controls.Add(this.ultraOptionSet1);
			this.ultraTabPageControl5.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl5.Controls.Add(this.label2);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(724, 335);
			this.ultraTabPageControl5.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl5_Paint);
			// 
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Banco"));
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsBanco;
			ultraGridColumn66.Header.VisiblePosition = 1;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn68.Header.VisiblePosition = 2;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 3;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 4;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 5;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 6;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 7;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn66,
																											ultraGridColumn67,
																											ultraGridColumn68,
																											ultraGridColumn69,
																											ultraGridColumn70,
																											ultraGridColumn71,
																											ultraGridColumn72,
																											ultraGridColumn73});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(112, 60);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(120, 21);
			this.cmbBanco.TabIndex = 37;
			this.cmbBanco.ValueMember = "Nombre";
			// 
			// cdsBanco
			// 
			this.cdsBanco.BindingContextCtrl = this;
			this.cdsBanco.DataLibrary = "LibContabilidad";
			this.cdsBanco.DataLibraryUrl = "";
			this.cdsBanco.DataSetDef = "dsCuentaCorriente";
			this.cdsBanco.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsBanco.Name = "cdsBanco";
			this.cdsBanco.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsBanco.SchemaDef = null;
			// 
			// btHoraExtra
			// 
			this.btHoraExtra.Location = new System.Drawing.Point(384, 198);
			this.btHoraExtra.Name = "btHoraExtra";
			this.btHoraExtra.Size = new System.Drawing.Size(72, 26);
			this.btHoraExtra.TabIndex = 36;
			this.btHoraExtra.Text = "Hora Extra";
			this.btHoraExtra.Click += new System.EventHandler(this.btHoraExtra_Click);
			// 
			// btHorario1
			// 
			this.btHorario1.Location = new System.Drawing.Point(24, 198);
			this.btHorario1.Name = "btHorario1";
			this.btHorario1.Size = new System.Drawing.Size(64, 26);
			this.btHorario1.TabIndex = 35;
			this.btHorario1.Text = "Horario";
			this.btHorario1.Click += new System.EventHandler(this.btHorario1_Click);
			// 
			// btDistribucion
			// 
			this.btDistribucion.Location = new System.Drawing.Point(336, 164);
			this.btDistribucion.Name = "btDistribucion";
			this.btDistribucion.Size = new System.Drawing.Size(24, 26);
			this.btDistribucion.TabIndex = 34;
			this.btDistribucion.Text = "...";
			this.btDistribucion.Click += new System.EventHandler(this.btDistribucion_Click);
			// 
			// btHorario
			// 
			this.btHorario.Location = new System.Drawing.Point(336, 198);
			this.btHorario.Name = "btHorario";
			this.btHorario.Size = new System.Drawing.Size(24, 26);
			this.btHorario.TabIndex = 33;
			this.btHorario.Text = "...";
			this.btHorario.Click += new System.EventHandler(this.btHorario_Click);
			// 
			// cmbCalendario
			// 
			this.cmbCalendario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCalendario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCalendario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idRolHorario"));
			this.cmbCalendario.DataMember = "RolHorario";
			this.cmbCalendario.DataSource = this.cdsSucursal;
			ultraGridBand13.AddButtonCaption = "RolTipo";
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Width = 106;
			ultraGridColumn75.Header.VisiblePosition = 0;
			ultraGridColumn75.Width = 326;
			ultraGridColumn76.Header.VisiblePosition = 2;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 17;
			ultraGridColumn77.Header.VisiblePosition = 3;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 17;
			ultraGridColumn78.Header.VisiblePosition = 4;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 17;
			ultraGridColumn79.Header.VisiblePosition = 5;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 17;
			ultraGridColumn80.Header.VisiblePosition = 6;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 17;
			ultraGridColumn81.Header.VisiblePosition = 7;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 17;
			ultraGridColumn82.Header.VisiblePosition = 8;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 17;
			ultraGridColumn83.Header.VisiblePosition = 9;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 17;
			ultraGridColumn84.Header.VisiblePosition = 10;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 17;
			ultraGridColumn85.Header.VisiblePosition = 11;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 17;
			ultraGridColumn86.Header.VisiblePosition = 12;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 17;
			ultraGridColumn87.Header.VisiblePosition = 13;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 17;
			ultraGridColumn88.Header.VisiblePosition = 14;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 17;
			ultraGridColumn89.Header.VisiblePosition = 15;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 17;
			ultraGridColumn90.Header.VisiblePosition = 16;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 17;
			ultraGridColumn91.Header.VisiblePosition = 17;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 17;
			ultraGridColumn92.Header.VisiblePosition = 18;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 17;
			ultraGridColumn93.Header.VisiblePosition = 19;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 17;
			ultraGridColumn94.Header.VisiblePosition = 20;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 17;
			ultraGridColumn95.Header.VisiblePosition = 21;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 17;
			ultraGridColumn96.Header.VisiblePosition = 22;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 17;
			ultraGridColumn97.Header.VisiblePosition = 23;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 17;
			ultraGridColumn98.Header.VisiblePosition = 24;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 17;
			ultraGridColumn99.Header.VisiblePosition = 25;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 17;
			ultraGridColumn100.Header.VisiblePosition = 26;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 17;
			ultraGridColumn101.Header.VisiblePosition = 27;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 17;
			ultraGridColumn102.Header.VisiblePosition = 28;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 17;
			ultraGridColumn103.Header.VisiblePosition = 29;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 17;
			ultraGridColumn104.Header.VisiblePosition = 30;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 17;
			ultraGridColumn105.Header.VisiblePosition = 31;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 17;
			ultraGridColumn106.Header.VisiblePosition = 32;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 17;
			ultraGridColumn107.Header.VisiblePosition = 33;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 17;
			ultraGridColumn108.Header.VisiblePosition = 34;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 17;
			ultraGridColumn109.Header.VisiblePosition = 35;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 17;
			ultraGridColumn110.Header.VisiblePosition = 36;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn110.Width = 17;
			ultraGridColumn111.Header.VisiblePosition = 37;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn111.Width = 17;
			ultraGridColumn112.Header.VisiblePosition = 38;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn112.Width = 17;
			ultraGridColumn113.Header.VisiblePosition = 39;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn113.Width = 17;
			ultraGridColumn114.Header.VisiblePosition = 40;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn114.Width = 17;
			ultraGridColumn115.Header.VisiblePosition = 41;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn115.Width = 17;
			ultraGridColumn116.Header.VisiblePosition = 42;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn116.Width = 17;
			ultraGridColumn117.Header.VisiblePosition = 43;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn117.Width = 17;
			ultraGridColumn118.Header.VisiblePosition = 44;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn118.Width = 17;
			ultraGridColumn119.Header.VisiblePosition = 45;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn119.Width = 17;
			ultraGridColumn120.Header.VisiblePosition = 46;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn120.Width = 17;
			ultraGridColumn121.Header.VisiblePosition = 47;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn121.Width = 17;
			ultraGridColumn122.Header.VisiblePosition = 48;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn122.Width = 17;
			ultraGridColumn123.Header.VisiblePosition = 49;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn123.Width = 17;
			ultraGridColumn124.Header.VisiblePosition = 50;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn124.Width = 17;
			ultraGridColumn125.Header.VisiblePosition = 51;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn125.Width = 17;
			ultraGridColumn126.Header.VisiblePosition = 52;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn126.Width = 17;
			ultraGridColumn127.Header.VisiblePosition = 53;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn127.Width = 17;
			ultraGridColumn128.Header.VisiblePosition = 54;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn128.Width = 17;
			ultraGridColumn129.Header.VisiblePosition = 55;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn129.Width = 17;
			ultraGridColumn130.Header.VisiblePosition = 56;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn130.Width = 17;
			ultraGridColumn131.Header.VisiblePosition = 57;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn131.Width = 17;
			ultraGridBand13.Columns.AddRange(new object[] {
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
																											ultraGridColumn105,
																											ultraGridColumn106,
																											ultraGridColumn107,
																											ultraGridColumn108,
																											ultraGridColumn109,
																											ultraGridColumn110,
																											ultraGridColumn111,
																											ultraGridColumn112,
																											ultraGridColumn113,
																											ultraGridColumn114,
																											ultraGridColumn115,
																											ultraGridColumn116,
																											ultraGridColumn117,
																											ultraGridColumn118,
																											ultraGridColumn119,
																											ultraGridColumn120,
																											ultraGridColumn121,
																											ultraGridColumn122,
																											ultraGridColumn123,
																											ultraGridColumn124,
																											ultraGridColumn125,
																											ultraGridColumn126,
																											ultraGridColumn127,
																											ultraGridColumn128,
																											ultraGridColumn129,
																											ultraGridColumn130,
																											ultraGridColumn131});
			this.cmbCalendario.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbCalendario.DisplayMember = "Nombre";
			this.cmbCalendario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCalendario.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCalendario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCalendario.Location = new System.Drawing.Point(112, 198);
			this.cmbCalendario.Name = "cmbCalendario";
			this.cmbCalendario.Size = new System.Drawing.Size(208, 21);
			this.cmbCalendario.TabIndex = 31;
			this.cmbCalendario.ValueMember = "idRolHorario";
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(24, 164);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(63, 16);
			this.label39.TabIndex = 30;
			this.label39.Text = "Distribucion";
			// 
			// cmbDistrib
			// 
			this.cmbDistrib.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDistrib.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDistrib.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idAsiDistGrupo"));
			this.cmbDistrib.DataMember = "AsientoDistGrupo";
			this.cmbDistrib.DataSource = this.cdsAsientoDistrib;
			ultraGridBand14.AddButtonCaption = "RolTipo";
			ultraGridColumn132.Header.VisiblePosition = 1;
			ultraGridColumn132.Width = 135;
			ultraGridColumn133.Header.VisiblePosition = 0;
			ultraGridColumn133.Width = 331;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn132,
																											ultraGridColumn133});
			this.cmbDistrib.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbDistrib.DisplayMember = "Distribucion";
			this.cmbDistrib.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDistrib.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDistrib.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDistrib.Location = new System.Drawing.Point(112, 164);
			this.cmbDistrib.Name = "cmbDistrib";
			this.cmbDistrib.Size = new System.Drawing.Size(208, 21);
			this.cmbDistrib.TabIndex = 29;
			this.cmbDistrib.ValueMember = "idAsiDistGrupo";
			this.cmbDistrib.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ultraCombo3_MouseDown);
			// 
			// cdsAsientoDistrib
			// 
			this.cdsAsientoDistrib.BindingContextCtrl = this;
			this.cdsAsientoDistrib.DataLibrary = "LibContabilidad";
			this.cdsAsientoDistrib.DataLibraryUrl = "";
			this.cdsAsientoDistrib.DataSetDef = "dsAsientoDist";
			this.cdsAsientoDistrib.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsientoDistrib.Name = "cdsAsientoDistrib";
			this.cdsAsientoDistrib.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsientoDistrib.SchemaDef = null;
			// 
			// ultraTextEditor4
			// 
			appearance30.BackColor = System.Drawing.Color.Gainsboro;
			appearance30.BackColor2 = System.Drawing.Color.White;
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor4.Appearance = appearance30;
			this.ultraTextEditor4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Estado_Civil"));
			this.ultraTextEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor4.Location = new System.Drawing.Point(112, 129);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(120, 22);
			this.ultraTextEditor4.TabIndex = 14;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(24, 129);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(61, 16);
			this.label26.TabIndex = 13;
			this.label26.Text = "Estado Civil";
			// 
			// ultraTextEditor3
			// 
			appearance31.BackColor = System.Drawing.Color.Gainsboro;
			appearance31.BackColor2 = System.Drawing.Color.White;
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor3.Appearance = appearance31;
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.CuentaCorr"));
			this.ultraTextEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor3.Location = new System.Drawing.Point(112, 95);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(120, 22);
			this.ultraTextEditor3.TabIndex = 12;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(24, 95);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(82, 16);
			this.label25.TabIndex = 11;
			this.label25.Text = "Numero Cuenta";
			// 
			// ultraTextEditor2
			// 
			appearance32.BackColor = System.Drawing.Color.Gainsboro;
			appearance32.BackColor2 = System.Drawing.Color.White;
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor2.Appearance = appearance32;
			this.ultraTextEditor2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Banco"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Location = new System.Drawing.Point(320, 129);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(152, 22);
			this.ultraTextEditor2.TabIndex = 10;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(24, 60);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(34, 16);
			this.label24.TabIndex = 9;
			this.label24.Text = "Banco";
			// 
			// ultraOptionSet1
			// 
			this.ultraOptionSet1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idTipoCuenta"));
			this.ultraOptionSet1.ItemAppearance = appearance33;
			this.ultraOptionSet1.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem7.DataValue = 1;
			valueListItem7.DisplayText = "Transf. Ahorros";
			valueListItem8.DataValue = 2;
			valueListItem8.DisplayText = "Transf. Corriente";
			valueListItem9.DataValue = 3;
			valueListItem9.DisplayText = "Transf. Virtual";
			valueListItem10.DataValue = 4;
			valueListItem10.DisplayText = "Cheque";
			this.ultraOptionSet1.Items.Add(valueListItem7);
			this.ultraOptionSet1.Items.Add(valueListItem8);
			this.ultraOptionSet1.Items.Add(valueListItem9);
			this.ultraOptionSet1.Items.Add(valueListItem10);
			this.ultraOptionSet1.ItemSpacingVertical = 6;
			this.ultraOptionSet1.Location = new System.Drawing.Point(320, 17);
			this.ultraOptionSet1.Name = "ultraOptionSet1";
			this.ultraOptionSet1.Size = new System.Drawing.Size(152, 104);
			this.ultraOptionSet1.TabIndex = 8;
			// 
			// ultraTextEditor1
			// 
			appearance34.BackColor = System.Drawing.Color.Gainsboro;
			appearance34.BackColor2 = System.Drawing.Color.White;
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor1.Appearance = appearance34;
			this.ultraTextEditor1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.CedulaMilitar"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(112, 26);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(120, 22);
			this.ultraTextEditor1.TabIndex = 7;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(24, 26);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Cédula Militar";
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.txtNumConadis);
			this.ultraTabPageControl6.Controls.Add(this.txtTipoDiscapacidad);
			this.ultraTabPageControl6.Controls.Add(this.chkVehiculo);
			this.ultraTabPageControl6.Controls.Add(this.label70);
			this.ultraTabPageControl6.Controls.Add(this.optVivienda);
			this.ultraTabPageControl6.Controls.Add(this.label69);
			this.ultraTabPageControl6.Controls.Add(this.txtPorcentajeDiscapacidad);
			this.ultraTabPageControl6.Controls.Add(this.chkDiscapacidad);
			this.ultraTabPageControl6.Controls.Add(this.label67);
			this.ultraTabPageControl6.Controls.Add(this.ultraCheckEditor3);
			this.ultraTabPageControl6.Controls.Add(this.ultraCheckEditor2);
			this.ultraTabPageControl6.Controls.Add(this.ultraCheckEditor1);
			this.ultraTabPageControl6.Controls.Add(this.ultraNumericEditor4);
			this.ultraTabPageControl6.Controls.Add(this.label37);
			this.ultraTabPageControl6.Controls.Add(this.ultraNumericEditor5);
			this.ultraTabPageControl6.Controls.Add(this.label38);
			this.ultraTabPageControl6.Controls.Add(this.ultraNumericEditor3);
			this.ultraTabPageControl6.Controls.Add(this.label36);
			this.ultraTabPageControl6.Controls.Add(this.ultraNumericEditor2);
			this.ultraTabPageControl6.Controls.Add(this.label35);
			this.ultraTabPageControl6.Controls.Add(this.label32);
			this.ultraTabPageControl6.Controls.Add(this.txtEmail);
			this.ultraTabPageControl6.Controls.Add(this.label31);
			this.ultraTabPageControl6.Controls.Add(this.ultraTextEditor8);
			this.ultraTabPageControl6.Controls.Add(this.label30);
			this.ultraTabPageControl6.Controls.Add(this.ultraTextEditor7);
			this.ultraTabPageControl6.Controls.Add(this.label29);
			this.ultraTabPageControl6.Controls.Add(this.ultraTextEditor6);
			this.ultraTabPageControl6.Controls.Add(this.label28);
			this.ultraTabPageControl6.Controls.Add(this.ultraTextEditor5);
			this.ultraTabPageControl6.Controls.Add(this.label68);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(724, 335);
			// 
			// txtNumConadis
			// 
			this.txtNumConadis.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumConadis.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.NumConadis"));
			this.txtNumConadis.Location = new System.Drawing.Point(608, 96);
			this.txtNumConadis.Name = "txtNumConadis";
			this.txtNumConadis.Size = new System.Drawing.Size(104, 22);
			this.txtNumConadis.TabIndex = 52;
			// 
			// txtTipoDiscapacidad
			// 
			this.txtTipoDiscapacidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTipoDiscapacidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.TipoDiscapacidad"));
			this.txtTipoDiscapacidad.Location = new System.Drawing.Point(608, 64);
			this.txtTipoDiscapacidad.Name = "txtTipoDiscapacidad";
			this.txtTipoDiscapacidad.Size = new System.Drawing.Size(104, 22);
			this.txtTipoDiscapacidad.TabIndex = 51;
			// 
			// chkVehiculo
			// 
			this.chkVehiculo.BackColor = System.Drawing.Color.Transparent;
			this.chkVehiculo.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Vehiculo"));
			this.chkVehiculo.Location = new System.Drawing.Point(512, 224);
			this.chkVehiculo.Name = "chkVehiculo";
			this.chkVehiculo.Size = new System.Drawing.Size(112, 21);
			this.chkVehiculo.TabIndex = 50;
			this.chkVehiculo.Text = "Vehiculo Propio";
			// 
			// label70
			// 
			this.label70.AutoSize = true;
			this.label70.BackColor = System.Drawing.Color.Transparent;
			this.label70.Location = new System.Drawing.Point(504, 136);
			this.label70.Name = "label70";
			this.label70.Size = new System.Drawing.Size(87, 16);
			this.label70.TabIndex = 49;
			this.label70.Text = "Tipo de Vivienda";
			// 
			// optVivienda
			// 
			this.optVivienda.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.TipoVivienda"));
			this.optVivienda.ItemAppearance = appearance35;
			this.optVivienda.ItemOrigin = new System.Drawing.Point(10, 2);
			valueListItem11.DataValue = 1;
			valueListItem11.DisplayText = "Propia";
			valueListItem12.DataValue = 2;
			valueListItem12.DisplayText = "Arrendada";
			valueListItem13.DataValue = 3;
			valueListItem13.DisplayText = "Familiar";
			this.optVivienda.Items.Add(valueListItem11);
			this.optVivienda.Items.Add(valueListItem12);
			this.optVivienda.Items.Add(valueListItem13);
			this.optVivienda.ItemSpacingHorizontal = 10;
			this.optVivienda.ItemSpacingVertical = 6;
			this.optVivienda.Location = new System.Drawing.Point(608, 136);
			this.optVivienda.Name = "optVivienda";
			this.optVivienda.Size = new System.Drawing.Size(96, 72);
			this.optVivienda.TabIndex = 48;
			// 
			// label69
			// 
			this.label69.AutoSize = true;
			this.label69.BackColor = System.Drawing.Color.Transparent;
			this.label69.Location = new System.Drawing.Point(504, 64);
			this.label69.Name = "label69";
			this.label69.Size = new System.Drawing.Size(94, 16);
			this.label69.TabIndex = 46;
			this.label69.Text = "Tipo Discapacidad";
			// 
			// txtPorcentajeDiscapacidad
			// 
			this.txtPorcentajeDiscapacidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.PorcDiscapacidad"));
			this.txtPorcentajeDiscapacidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentajeDiscapacidad.Location = new System.Drawing.Point(648, 24);
			this.txtPorcentajeDiscapacidad.MaxValue = 100;
			this.txtPorcentajeDiscapacidad.MinValue = 0;
			this.txtPorcentajeDiscapacidad.Name = "txtPorcentajeDiscapacidad";
			this.txtPorcentajeDiscapacidad.PromptChar = ' ';
			this.txtPorcentajeDiscapacidad.Size = new System.Drawing.Size(64, 22);
			this.txtPorcentajeDiscapacidad.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPorcentajeDiscapacidad.TabIndex = 45;
			// 
			// chkDiscapacidad
			// 
			this.chkDiscapacidad.BackColor = System.Drawing.Color.Transparent;
			this.chkDiscapacidad.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.PoseeDiscapacidad"));
			this.chkDiscapacidad.Location = new System.Drawing.Point(504, 24);
			this.chkDiscapacidad.Name = "chkDiscapacidad";
			this.chkDiscapacidad.Size = new System.Drawing.Size(88, 21);
			this.chkDiscapacidad.TabIndex = 44;
			this.chkDiscapacidad.Text = "Discapacidad";
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Location = new System.Drawing.Point(504, 96);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(74, 16);
			this.label67.TabIndex = 42;
			this.label67.Text = "Num. Conadis";
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Sindicato"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(304, 24);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(136, 21);
			this.ultraCheckEditor3.TabIndex = 41;
			this.ultraCheckEditor3.Text = "Sindicato";
			// 
			// ultraCheckEditor2
			// 
			this.ultraCheckEditor2.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.Asociacion"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(304, 56);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(136, 22);
			this.ultraCheckEditor2.TabIndex = 40;
			this.ultraCheckEditor2.Text = "Asociacion";
			// 
			// ultraCheckEditor1
			// 
			this.ultraCheckEditor1.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.TieneQuincena"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(304, 168);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(136, 21);
			this.ultraCheckEditor1.TabIndex = 39;
			this.ultraCheckEditor1.Text = "Quincena";
			// 
			// ultraNumericEditor4
			// 
			this.ultraNumericEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Desc2"));
			this.ultraNumericEditor4.FormatString = "#,##0.00";
			this.ultraNumericEditor4.Location = new System.Drawing.Point(344, 232);
			this.ultraNumericEditor4.Name = "ultraNumericEditor4";
			this.ultraNumericEditor4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor4.PromptChar = ' ';
			this.ultraNumericEditor4.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor4.TabIndex = 38;
			this.ultraNumericEditor4.ValueChanged += new System.EventHandler(this.ultraNumericEditor4_ValueChanged);
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(248, 232);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(70, 16);
			this.label37.TabIndex = 37;
			this.label37.Text = "Descuento 2:";
			this.label37.Click += new System.EventHandler(this.label37_Click);
			// 
			// ultraNumericEditor5
			// 
			this.ultraNumericEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Desc1"));
			this.ultraNumericEditor5.FormatString = "#,##0.00";
			this.ultraNumericEditor5.Location = new System.Drawing.Point(128, 233);
			this.ultraNumericEditor5.Name = "ultraNumericEditor5";
			this.ultraNumericEditor5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor5.PromptChar = ' ';
			this.ultraNumericEditor5.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor5.TabIndex = 36;
			this.ultraNumericEditor5.ValueChanged += new System.EventHandler(this.ultraNumericEditor5_ValueChanged);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(24, 233);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(66, 16);
			this.label38.TabIndex = 35;
			this.label38.Text = "Descuento 1";
			this.label38.Click += new System.EventHandler(this.label38_Click);
			// 
			// ultraNumericEditor3
			// 
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Bono1"));
			this.ultraNumericEditor3.FormatString = "#,##0.00";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(344, 200);
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor3.TabIndex = 34;
			this.ultraNumericEditor3.ValueChanged += new System.EventHandler(this.ultraNumericEditor3_ValueChanged);
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(248, 200);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(43, 16);
			this.label36.TabIndex = 33;
			this.label36.Text = "Bono 1:";
			this.label36.Click += new System.EventHandler(this.label36_Click);
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Bono"));
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(128, 198);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor2.TabIndex = 32;
			this.ultraNumericEditor2.ValueChanged += new System.EventHandler(this.ultraNumericEditor2_ValueChanged);
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(24, 198);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(33, 16);
			this.label35.TabIndex = 31;
			this.label35.Text = "Bono:";
			this.label35.Click += new System.EventHandler(this.label35_Click);
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Location = new System.Drawing.Point(24, 164);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(30, 16);
			this.label32.TabIndex = 9;
			this.label32.Text = "EMail";
			// 
			// txtEmail
			// 
			this.txtEmail.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.email"));
			this.txtEmail.Location = new System.Drawing.Point(128, 164);
			this.txtEmail.Name = "txtEmail";
			this.txtEmail.Size = new System.Drawing.Size(160, 22);
			this.txtEmail.TabIndex = 8;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(24, 129);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(78, 16);
			this.label31.TabIndex = 7;
			this.label31.Text = "Tarjeta Oficina";
			this.label31.Click += new System.EventHandler(this.label31_Click);
			// 
			// ultraTextEditor8
			// 
			this.ultraTextEditor8.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.TarjetaOficina"));
			this.ultraTextEditor8.Location = new System.Drawing.Point(128, 129);
			this.ultraTextEditor8.Name = "ultraTextEditor8";
			this.ultraTextEditor8.Size = new System.Drawing.Size(160, 22);
			this.ultraTextEditor8.TabIndex = 6;
			this.ultraTextEditor8.ValueChanged += new System.EventHandler(this.ultraTextEditor8_ValueChanged);
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(24, 95);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(79, 16);
			this.label30.TabIndex = 5;
			this.label30.Text = "Tarjeta Edificio";
			this.label30.Click += new System.EventHandler(this.label30_Click);
			// 
			// ultraTextEditor7
			// 
			this.ultraTextEditor7.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.TarjEdificio"));
			this.ultraTextEditor7.Location = new System.Drawing.Point(128, 95);
			this.ultraTextEditor7.Name = "ultraTextEditor7";
			this.ultraTextEditor7.Size = new System.Drawing.Size(160, 22);
			this.ultraTextEditor7.TabIndex = 4;
			this.ultraTextEditor7.ValueChanged += new System.EventHandler(this.ultraTextEditor7_ValueChanged);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(24, 60);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(93, 16);
			this.label29.TabIndex = 3;
			this.label29.Text = "Codigo Telefonico";
			this.label29.Click += new System.EventHandler(this.label29_Click);
			// 
			// ultraTextEditor6
			// 
			this.ultraTextEditor6.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.CodTelefono"));
			this.ultraTextEditor6.Location = new System.Drawing.Point(128, 60);
			this.ultraTextEditor6.Name = "ultraTextEditor6";
			this.ultraTextEditor6.Size = new System.Drawing.Size(160, 22);
			this.ultraTextEditor6.TabIndex = 2;
			this.ultraTextEditor6.ValueChanged += new System.EventHandler(this.ultraTextEditor6_ValueChanged);
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(24, 26);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(80, 16);
			this.label28.TabIndex = 1;
			this.label28.Text = "Codigo Ingreso";
			this.label28.Click += new System.EventHandler(this.label28_Click);
			// 
			// ultraTextEditor5
			// 
			this.ultraTextEditor5.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.CodIngreso"));
			this.ultraTextEditor5.Location = new System.Drawing.Point(128, 26);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(160, 22);
			this.ultraTextEditor5.TabIndex = 0;
			this.ultraTextEditor5.ValueChanged += new System.EventHandler(this.ultraTextEditor5_ValueChanged);
			// 
			// label68
			// 
			this.label68.AutoSize = true;
			this.label68.BackColor = System.Drawing.Color.Transparent;
			this.label68.Location = new System.Drawing.Point(624, 24);
			this.label68.Name = "label68";
			this.label68.Size = new System.Drawing.Size(15, 16);
			this.label68.TabIndex = 42;
			this.label68.Text = "%";
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.cmbTitulo);
			this.ultraTabPageControl7.Controls.Add(this.cmbNivel);
			this.ultraTabPageControl7.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(724, 335);
			// 
			// cmbTitulo
			// 
			this.cmbTitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTitulo.DataMember = "PersTitulo";
			this.cmbTitulo.DataSource = this.cdsProvincia;
			ultraGridColumn134.Header.VisiblePosition = 1;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 0;
			ultraGridColumn135.Width = 163;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn134,
																											ultraGridColumn135});
			this.cmbTitulo.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbTitulo.DisplayMember = "Titulo";
			this.cmbTitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTitulo.Location = new System.Drawing.Point(288, 95);
			this.cmbTitulo.Name = "cmbTitulo";
			this.cmbTitulo.Size = new System.Drawing.Size(184, 95);
			this.cmbTitulo.TabIndex = 2;
			this.cmbTitulo.ValueMember = "idPersTitulo";
			this.cmbTitulo.Visible = false;
			// 
			// cmbNivel
			// 
			this.cmbNivel.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNivel.DataMember = "PersNivel";
			this.cmbNivel.DataSource = this.cdsProvincia;
			ultraGridColumn136.Header.VisiblePosition = 1;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 0;
			ultraGridColumn137.Width = 172;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn136,
																											ultraGridColumn137});
			this.cmbNivel.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbNivel.DisplayMember = "Nivel";
			this.cmbNivel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNivel.Location = new System.Drawing.Point(56, 95);
			this.cmbNivel.Name = "cmbNivel";
			this.cmbNivel.Size = new System.Drawing.Size(192, 95);
			this.cmbNivel.TabIndex = 1;
			this.cmbNivel.ValueMember = "idNivel";
			this.cmbNivel.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "PersEstudios";
			this.ultraGrid1.DataSource = this.cdsProvincia;
			appearance36.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance36;
			ultraGridColumn138.Header.VisiblePosition = 4;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 5;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.Caption = "Nivel";
			ultraGridColumn140.Header.VisiblePosition = 0;
			ultraGridColumn140.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn140.Width = 110;
			ultraGridColumn141.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn141.Header.VisiblePosition = 1;
			ultraGridColumn141.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Edit;
			ultraGridColumn141.Width = 136;
			ultraGridColumn142.Header.Caption = "Titulo";
			ultraGridColumn142.Header.VisiblePosition = 2;
			ultraGridColumn142.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn142.Width = 106;
			ultraGridColumn143.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn143.Header.VisiblePosition = 3;
			ultraGridColumn143.Width = 139;
			ultraGridColumn144.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn144.Header.Caption = "FechaDesde";
			ultraGridColumn144.Header.VisiblePosition = 6;
			ultraGridColumn145.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn145.Header.VisiblePosition = 7;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn138,
																											ultraGridColumn139,
																											ultraGridColumn140,
																											ultraGridColumn141,
																											ultraGridColumn142,
																											ultraGridColumn143,
																											ultraGridColumn144,
																											ultraGridColumn145});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance38.FontData.BoldAsString = "True";
			appearance38.FontData.Name = "Arial";
			appearance38.FontData.SizeInPoints = 10F;
			appearance38.ForeColor = System.Drawing.Color.White;
			appearance38.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance40;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(712, 320);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// ultraTabPageControl8
			// 
			this.ultraTabPageControl8.Controls.Add(this.cmbParentesco);
			this.ultraTabPageControl8.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl8.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl8.Name = "ultraTabPageControl8";
			this.ultraTabPageControl8.Size = new System.Drawing.Size(724, 335);
			// 
			// cmbParentesco
			// 
			this.cmbParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParentesco.DataMember = "PersParentesco";
			this.cmbParentesco.DataSource = this.cdsProvincia;
			ultraGridColumn146.Header.VisiblePosition = 1;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 0;
			ultraGridColumn147.Width = 211;
			ultraGridColumn148.Header.VisiblePosition = 2;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn146,
																											ultraGridColumn147,
																											ultraGridColumn148});
			this.cmbParentesco.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.cmbParentesco.DisplayMember = "Parentesco";
			this.cmbParentesco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParentesco.Location = new System.Drawing.Point(248, 121);
			this.cmbParentesco.Name = "cmbParentesco";
			this.cmbParentesco.Size = new System.Drawing.Size(232, 86);
			this.cmbParentesco.TabIndex = 1;
			this.cmbParentesco.ValueMember = "idParentesco";
			this.cmbParentesco.Visible = false;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "PersFamiliar";
			this.ultraGrid2.DataSource = this.cdsProvincia;
			appearance41.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance41;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn149.Header.VisiblePosition = 4;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 5;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.Caption = "Parentesco";
			ultraGridColumn151.Header.VisiblePosition = 1;
			ultraGridColumn151.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn151.Width = 197;
			ultraGridColumn152.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn152.Header.VisiblePosition = 0;
			ultraGridColumn152.Width = 213;
			ultraGridColumn153.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn153.Header.VisiblePosition = 2;
			ultraGridColumn153.Width = 141;
			ultraGridColumn154.Header.Caption = "Nacimiento";
			ultraGridColumn154.Header.VisiblePosition = 3;
			ultraGridColumn154.Width = 116;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn149,
																											ultraGridColumn150,
																											ultraGridColumn151,
																											ultraGridColumn152,
																											ultraGridColumn153,
																											ultraGridColumn154});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance42.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance43.FontData.BoldAsString = "True";
			appearance43.FontData.Name = "Arial";
			appearance43.FontData.SizeInPoints = 10F;
			appearance43.ForeColor = System.Drawing.Color.White;
			appearance43.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(16, 17);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(688, 233);
			this.ultraGrid2.TabIndex = 0;
			this.ultraGrid2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid2_AfterRowInsert);
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.btRepVacacion);
			this.ultraTabPageControl9.Controls.Add(this.cmbHastaV);
			this.ultraTabPageControl9.Controls.Add(this.ultraLabel2);
			this.ultraTabPageControl9.Controls.Add(this.ultraLabel1);
			this.ultraTabPageControl9.Controls.Add(this.cmbDesdeV);
			this.ultraTabPageControl9.Controls.Add(this.ultraGrid3);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(724, 335);
			// 
			// btRepVacacion
			// 
			this.btRepVacacion.Location = new System.Drawing.Point(368, 233);
			this.btRepVacacion.Name = "btRepVacacion";
			this.btRepVacacion.Size = new System.Drawing.Size(64, 25);
			this.btRepVacacion.TabIndex = 6;
			this.btRepVacacion.Text = "&Reporte";
			this.btRepVacacion.Click += new System.EventHandler(this.btRepVacacion_Click);
			this.btRepVacacion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btRepVacacion_MouseUp);
			// 
			// cmbHastaV
			// 
			this.cmbHastaV.DateTime = new System.DateTime(2011, 5, 4, 0, 0, 0, 0);
			this.cmbHastaV.FormatString = "dd/MMM/yyyy";
			this.cmbHastaV.Location = new System.Drawing.Point(240, 233);
			this.cmbHastaV.Name = "cmbHastaV";
			this.cmbHastaV.Size = new System.Drawing.Size(112, 22);
			this.cmbHastaV.TabIndex = 5;
			this.cmbHastaV.Value = new System.DateTime(2011, 5, 4, 0, 0, 0, 0);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Location = new System.Drawing.Point(192, 233);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(33, 15);
			this.ultraLabel2.TabIndex = 4;
			this.ultraLabel2.Text = "Hasta";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 233);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel1.TabIndex = 3;
			this.ultraLabel1.Text = "Desde";
			// 
			// cmbDesdeV
			// 
			this.cmbDesdeV.DateTime = new System.DateTime(2011, 5, 4, 0, 0, 0, 0);
			this.cmbDesdeV.FormatString = "dd/MMM/yyyy";
			this.cmbDesdeV.Location = new System.Drawing.Point(64, 233);
			this.cmbDesdeV.Name = "cmbDesdeV";
			this.cmbDesdeV.Size = new System.Drawing.Size(112, 22);
			this.cmbDesdeV.TabIndex = 2;
			this.cmbDesdeV.Value = new System.DateTime(2011, 5, 4, 0, 0, 0, 0);
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid3.DataMember = "RolVacacion";
			this.ultraGrid3.DataSource = this.cdsProvincia;
			appearance46.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance46;
			this.ultraGrid3.DisplayLayout.AutoFitColumns = true;
			ultraGridBand20.AddButtonCaption = "PersFamiliar";
			ultraGridColumn155.Header.VisiblePosition = 0;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn155.Width = 131;
			ultraGridColumn156.Format = "dd/MMM/yyyy";
			ultraGridColumn156.Header.VisiblePosition = 1;
			ultraGridColumn156.Width = 202;
			ultraGridColumn157.Header.VisiblePosition = 2;
			ultraGridColumn157.Width = 48;
			ultraGridColumn158.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn158.Header.VisiblePosition = 4;
			ultraGridColumn158.Width = 369;
			ultraGridColumn159.Header.VisiblePosition = 5;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn159.Width = 63;
			ultraGridColumn160.Header.VisiblePosition = 3;
			ultraGridColumn160.Width = 56;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn155,
																											ultraGridColumn156,
																											ultraGridColumn157,
																											ultraGridColumn158,
																											ultraGridColumn159,
																											ultraGridColumn160});
			this.ultraGrid3.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.ultraGrid3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance48.FontData.BoldAsString = "True";
			appearance48.FontData.Name = "Arial";
			appearance48.FontData.SizeInPoints = 10F;
			appearance48.ForeColor = System.Drawing.Color.White;
			appearance48.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance50;
			this.ultraGrid3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid3.Location = new System.Drawing.Point(16, 16);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(696, 208);
			this.ultraGrid3.TabIndex = 1;
			this.ultraGrid3.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid3_AfterRowInsert);
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.ultraGrid4);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(724, 335);
			// 
			// ultraGrid4
			// 
			this.ultraGrid4.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid4.DataMember = "RolSueldo";
			this.ultraGrid4.DataSource = this.cdsProvincia;
			appearance51.BackColor = System.Drawing.Color.White;
			this.ultraGrid4.DisplayLayout.Appearance = appearance51;
			this.ultraGrid4.DisplayLayout.AutoFitColumns = true;
			ultraGridBand21.AddButtonCaption = "PersFamiliar";
			ultraGridColumn161.Header.VisiblePosition = 0;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn161.Width = 55;
			ultraGridColumn162.Header.VisiblePosition = 1;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn162.Width = 74;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn163.CellAppearance = appearance52;
			ultraGridColumn163.Format = "#,##0.00";
			ultraGridColumn163.Header.VisiblePosition = 3;
			ultraGridColumn163.Width = 88;
			ultraGridColumn164.Format = "dd/MMM/yyyy";
			ultraGridColumn164.Header.VisiblePosition = 2;
			ultraGridColumn164.Width = 217;
			ultraGridColumn165.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn165.Header.VisiblePosition = 4;
			ultraGridColumn165.Width = 352;
			ultraGridColumn166.Header.VisiblePosition = 5;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn166.Width = 64;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn161,
																											ultraGridColumn162,
																											ultraGridColumn163,
																											ultraGridColumn164,
																											ultraGridColumn165,
																											ultraGridColumn166});
			this.ultraGrid4.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.ultraGrid4.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid4.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid4.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance53.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid4.DisplayLayout.Override.CardAreaAppearance = appearance53;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance54.FontData.BoldAsString = "True";
			appearance54.FontData.Name = "Arial";
			appearance54.FontData.SizeInPoints = 10F;
			appearance54.ForeColor = System.Drawing.Color.White;
			appearance54.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid4.DisplayLayout.Override.HeaderAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.RowSelectorAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid4.DisplayLayout.Override.SelectedRowAppearance = appearance56;
			this.ultraGrid4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid4.Location = new System.Drawing.Point(18, 16);
			this.ultraGrid4.Name = "ultraGrid4";
			this.ultraGrid4.Size = new System.Drawing.Size(678, 233);
			this.ultraGrid4.TabIndex = 2;
			this.ultraGrid4.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid4_AfterRowInsert);
			// 
			// ultraTabPageControl11
			// 
			this.ultraTabPageControl11.Controls.Add(this.ultraCheckEditor4);
			this.ultraTabPageControl11.Controls.Add(this.ultraNumericEditor14);
			this.ultraTabPageControl11.Controls.Add(this.label56);
			this.ultraTabPageControl11.Controls.Add(this.ultraNumericEditor11);
			this.ultraTabPageControl11.Controls.Add(this.label53);
			this.ultraTabPageControl11.Controls.Add(this.ultraNumericEditor12);
			this.ultraTabPageControl11.Controls.Add(this.label54);
			this.ultraTabPageControl11.Controls.Add(this.ultraNumericEditor13);
			this.ultraTabPageControl11.Controls.Add(this.label55);
			this.ultraTabPageControl11.Controls.Add(this.ultraNumericEditor10);
			this.ultraTabPageControl11.Controls.Add(this.label52);
			this.ultraTabPageControl11.Controls.Add(this.ultraNumericEditor9);
			this.ultraTabPageControl11.Controls.Add(this.label51);
			this.ultraTabPageControl11.Controls.Add(this.btTotalGasto);
			this.ultraTabPageControl11.Controls.Add(this.txtTotalGasto);
			this.ultraTabPageControl11.Controls.Add(this.groupBox1);
			this.ultraTabPageControl11.Controls.Add(this.groupBox2);
			this.ultraTabPageControl11.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl11.Name = "ultraTabPageControl11";
			this.ultraTabPageControl11.Size = new System.Drawing.Size(724, 335);
			// 
			// ultraCheckEditor4
			// 
			this.ultraCheckEditor4.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor4.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPersonal, "Personal.GTC"));
			this.ultraCheckEditor4.Location = new System.Drawing.Point(16, 288);
			this.ultraCheckEditor4.Name = "ultraCheckEditor4";
			this.ultraCheckEditor4.Size = new System.Drawing.Size(176, 21);
			this.ultraCheckEditor4.TabIndex = 75;
			this.ultraCheckEditor4.Text = "Tarjeta de Crédito Corporativa";
			// 
			// ultraNumericEditor14
			// 
			this.ultraNumericEditor14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.ImpRentaAsumido"));
			this.ultraNumericEditor14.FormatString = "#,##0.00";
			this.ultraNumericEditor14.Location = new System.Drawing.Point(600, 200);
			this.ultraNumericEditor14.Name = "ultraNumericEditor14";
			this.ultraNumericEditor14.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor14.PromptChar = ' ';
			this.ultraNumericEditor14.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor14.TabIndex = 67;
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.BackColor = System.Drawing.Color.Transparent;
			this.label56.Location = new System.Drawing.Point(496, 200);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(106, 16);
			this.label56.TabIndex = 66;
			this.label56.Text = "Imp. Renta Asumido";
			this.label56.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// ultraNumericEditor11
			// 
			this.ultraNumericEditor11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gRebajaIIIEdad"));
			this.ultraNumericEditor11.FormatString = "#,##0.00";
			this.ultraNumericEditor11.Location = new System.Drawing.Point(600, 168);
			this.ultraNumericEditor11.Name = "ultraNumericEditor11";
			this.ultraNumericEditor11.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor11.PromptChar = ' ';
			this.ultraNumericEditor11.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor11.TabIndex = 59;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(496, 168);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(83, 16);
			this.label53.TabIndex = 58;
			this.label53.Text = "Rebaja III Edad";
			// 
			// ultraNumericEditor12
			// 
			this.ultraNumericEditor12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gDiscapacidad"));
			this.ultraNumericEditor12.FormatString = "#,##0.00";
			this.ultraNumericEditor12.Location = new System.Drawing.Point(600, 136);
			this.ultraNumericEditor12.Name = "ultraNumericEditor12";
			this.ultraNumericEditor12.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor12.PromptChar = ' ';
			this.ultraNumericEditor12.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor12.TabIndex = 57;
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Location = new System.Drawing.Point(496, 136);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(69, 16);
			this.label54.TabIndex = 56;
			this.label54.Text = "Discapacidad";
			this.label54.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// ultraNumericEditor13
			// 
			this.ultraNumericEditor13.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gDesahucio"));
			this.ultraNumericEditor13.FormatString = "#,##0.00";
			this.ultraNumericEditor13.Location = new System.Drawing.Point(600, 96);
			this.ultraNumericEditor13.Name = "ultraNumericEditor13";
			this.ultraNumericEditor13.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor13.PromptChar = ' ';
			this.ultraNumericEditor13.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor13.TabIndex = 55;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.BackColor = System.Drawing.Color.Transparent;
			this.label55.Location = new System.Drawing.Point(496, 96);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(55, 16);
			this.label55.TabIndex = 54;
			this.label55.Text = "Desahucio";
			// 
			// ultraNumericEditor10
			// 
			this.ultraNumericEditor10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.PorcentajeComision"));
			this.ultraNumericEditor10.FormatString = "#,##0.00";
			this.ultraNumericEditor10.Location = new System.Drawing.Point(600, 64);
			this.ultraNumericEditor10.Name = "ultraNumericEditor10";
			this.ultraNumericEditor10.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor10.PromptChar = ' ';
			this.ultraNumericEditor10.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor10.TabIndex = 53;
			// 
			// label52
			// 
			this.label52.AutoSize = true;
			this.label52.BackColor = System.Drawing.Color.Transparent;
			this.label52.Location = new System.Drawing.Point(496, 64);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(64, 16);
			this.label52.TabIndex = 52;
			this.label52.Text = "% Comision";
			// 
			// ultraNumericEditor9
			// 
			this.ultraNumericEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.BaseComision"));
			this.ultraNumericEditor9.FormatString = "#,##0.00";
			this.ultraNumericEditor9.Location = new System.Drawing.Point(600, 32);
			this.ultraNumericEditor9.Name = "ultraNumericEditor9";
			this.ultraNumericEditor9.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor9.PromptChar = ' ';
			this.ultraNumericEditor9.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor9.TabIndex = 51;
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.BackColor = System.Drawing.Color.Transparent;
			this.label51.Location = new System.Drawing.Point(496, 32);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(77, 16);
			this.label51.TabIndex = 50;
			this.label51.Text = "Base Comision";
			// 
			// btTotalGasto
			// 
			this.btTotalGasto.Location = new System.Drawing.Point(24, 240);
			this.btTotalGasto.Name = "btTotalGasto";
			this.btTotalGasto.Size = new System.Drawing.Size(88, 26);
			this.btTotalGasto.TabIndex = 49;
			this.btTotalGasto.Text = "Gasto Mes";
			this.btTotalGasto.Click += new System.EventHandler(this.btTotalGasto_Click);
			// 
			// txtTotalGasto
			// 
			this.txtTotalGasto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gTotal"));
			this.txtTotalGasto.FormatString = "#,##0.00";
			this.txtTotalGasto.Location = new System.Drawing.Point(136, 240);
			this.txtTotalGasto.Name = "txtTotalGasto";
			this.txtTotalGasto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGasto.PromptChar = ' ';
			this.txtTotalGasto.Size = new System.Drawing.Size(100, 22);
			this.txtTotalGasto.TabIndex = 48;
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.Color.Transparent;
			this.groupBox1.Controls.Add(this.ultraNumericEditor18);
			this.groupBox1.Controls.Add(this.label60);
			this.groupBox1.Controls.Add(this.ultraNumericEditor15);
			this.groupBox1.Controls.Add(this.label57);
			this.groupBox1.Controls.Add(this.ultraNumericEditor16);
			this.groupBox1.Controls.Add(this.label58);
			this.groupBox1.Controls.Add(this.ultraNumericEditor17);
			this.groupBox1.Controls.Add(this.label59);
			this.groupBox1.Location = new System.Drawing.Point(280, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(200, 181);
			this.groupBox1.TabIndex = 68;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Otro Empleador";
			// 
			// ultraNumericEditor18
			// 
			this.ultraNumericEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.RentaRetOE"));
			this.ultraNumericEditor18.FormatString = "#,##0.00";
			this.ultraNumericEditor18.Location = new System.Drawing.Point(88, 138);
			this.ultraNumericEditor18.Name = "ultraNumericEditor18";
			this.ultraNumericEditor18.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor18.PromptChar = ' ';
			this.ultraNumericEditor18.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor18.TabIndex = 73;
			// 
			// label60
			// 
			this.label60.AutoSize = true;
			this.label60.BackColor = System.Drawing.Color.Transparent;
			this.label60.Location = new System.Drawing.Point(8, 138);
			this.label60.Name = "label60";
			this.label60.Size = new System.Drawing.Size(57, 16);
			this.label60.TabIndex = 72;
			this.label60.Text = "Renta Ret.";
			// 
			// ultraNumericEditor15
			// 
			this.ultraNumericEditor15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.DeduccionOE"));
			this.ultraNumericEditor15.FormatString = "#,##0.00";
			this.ultraNumericEditor15.Location = new System.Drawing.Point(88, 69);
			this.ultraNumericEditor15.Name = "ultraNumericEditor15";
			this.ultraNumericEditor15.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor15.PromptChar = ' ';
			this.ultraNumericEditor15.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor15.TabIndex = 71;
			// 
			// label57
			// 
			this.label57.AutoSize = true;
			this.label57.BackColor = System.Drawing.Color.Transparent;
			this.label57.Location = new System.Drawing.Point(8, 69);
			this.label57.Name = "label57";
			this.label57.Size = new System.Drawing.Size(56, 16);
			this.label57.TabIndex = 70;
			this.label57.Text = "Deduccion";
			// 
			// ultraNumericEditor16
			// 
			this.ultraNumericEditor16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.OtrasRebOE"));
			this.ultraNumericEditor16.FormatString = "#,##0.00";
			this.ultraNumericEditor16.Location = new System.Drawing.Point(88, 103);
			this.ultraNumericEditor16.Name = "ultraNumericEditor16";
			this.ultraNumericEditor16.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor16.PromptChar = ' ';
			this.ultraNumericEditor16.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor16.TabIndex = 69;
			// 
			// label58
			// 
			this.label58.AutoSize = true;
			this.label58.BackColor = System.Drawing.Color.Transparent;
			this.label58.Location = new System.Drawing.Point(8, 103);
			this.label58.Name = "label58";
			this.label58.Size = new System.Drawing.Size(74, 16);
			this.label58.TabIndex = 68;
			this.label58.Text = "Otras Rebajas";
			// 
			// ultraNumericEditor17
			// 
			this.ultraNumericEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.IngGravOE"));
			this.ultraNumericEditor17.FormatString = "#,##0.00";
			this.ultraNumericEditor17.Location = new System.Drawing.Point(88, 34);
			this.ultraNumericEditor17.Name = "ultraNumericEditor17";
			this.ultraNumericEditor17.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor17.PromptChar = ' ';
			this.ultraNumericEditor17.Size = new System.Drawing.Size(100, 22);
			this.ultraNumericEditor17.TabIndex = 67;
			// 
			// label59
			// 
			this.label59.AutoSize = true;
			this.label59.BackColor = System.Drawing.Color.Transparent;
			this.label59.Location = new System.Drawing.Point(8, 34);
			this.label59.Name = "label59";
			this.label59.Size = new System.Drawing.Size(55, 16);
			this.label59.TabIndex = 66;
			this.label59.Text = "Ing. Grab.";
			// 
			// groupBox2
			// 
			this.groupBox2.BackColor = System.Drawing.Color.Transparent;
			this.groupBox2.Controls.Add(this.txtVestimenta);
			this.groupBox2.Controls.Add(this.label50);
			this.groupBox2.Controls.Add(this.txtEducacion);
			this.groupBox2.Controls.Add(this.label49);
			this.groupBox2.Controls.Add(this.txtSalud);
			this.groupBox2.Controls.Add(this.label48);
			this.groupBox2.Controls.Add(this.txtAlimentacion);
			this.groupBox2.Controls.Add(this.label47);
			this.groupBox2.Controls.Add(this.txtVivienda);
			this.groupBox2.Controls.Add(this.label46);
			this.groupBox2.Location = new System.Drawing.Point(16, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(240, 208);
			this.groupBox2.TabIndex = 74;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Gastos Anuales";
			// 
			// txtVestimenta
			// 
			this.txtVestimenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gVestimenta"));
			this.txtVestimenta.FormatString = "#,##0.00";
			this.txtVestimenta.Location = new System.Drawing.Point(120, 168);
			this.txtVestimenta.Name = "txtVestimenta";
			this.txtVestimenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVestimenta.PromptChar = ' ';
			this.txtVestimenta.Size = new System.Drawing.Size(100, 22);
			this.txtVestimenta.TabIndex = 56;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.BackColor = System.Drawing.Color.Transparent;
			this.label50.Location = new System.Drawing.Point(16, 168);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(60, 16);
			this.label50.TabIndex = 55;
			this.label50.Text = "Vestimenta";
			// 
			// txtEducacion
			// 
			this.txtEducacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gEducacion"));
			this.txtEducacion.FormatString = "#,##0.00";
			this.txtEducacion.Location = new System.Drawing.Point(120, 128);
			this.txtEducacion.Name = "txtEducacion";
			this.txtEducacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEducacion.PromptChar = ' ';
			this.txtEducacion.Size = new System.Drawing.Size(100, 22);
			this.txtEducacion.TabIndex = 54;
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.BackColor = System.Drawing.Color.Transparent;
			this.label49.Location = new System.Drawing.Point(16, 128);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(55, 16);
			this.label49.TabIndex = 53;
			this.label49.Text = "Educacion";
			this.label49.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// txtSalud
			// 
			this.txtSalud.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gSalud"));
			this.txtSalud.FormatString = "#,##0.00";
			this.txtSalud.Location = new System.Drawing.Point(120, 64);
			this.txtSalud.Name = "txtSalud";
			this.txtSalud.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSalud.PromptChar = ' ';
			this.txtSalud.Size = new System.Drawing.Size(100, 22);
			this.txtSalud.TabIndex = 52;
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.BackColor = System.Drawing.Color.Transparent;
			this.label48.Location = new System.Drawing.Point(16, 64);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(32, 16);
			this.label48.TabIndex = 51;
			this.label48.Text = "Salud";
			// 
			// txtAlimentacion
			// 
			this.txtAlimentacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gAlimentacion"));
			this.txtAlimentacion.FormatString = "#,##0.00";
			this.txtAlimentacion.Location = new System.Drawing.Point(120, 96);
			this.txtAlimentacion.Name = "txtAlimentacion";
			this.txtAlimentacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAlimentacion.PromptChar = ' ';
			this.txtAlimentacion.Size = new System.Drawing.Size(100, 22);
			this.txtAlimentacion.TabIndex = 50;
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(16, 96);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(68, 16);
			this.label47.TabIndex = 49;
			this.label47.Text = "Alimentacion";
			// 
			// txtVivienda
			// 
			this.txtVivienda.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.gVivienda"));
			this.txtVivienda.FormatString = "#,##0.00";
			this.txtVivienda.Location = new System.Drawing.Point(120, 24);
			this.txtVivienda.Name = "txtVivienda";
			this.txtVivienda.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVivienda.PromptChar = ' ';
			this.txtVivienda.Size = new System.Drawing.Size(100, 22);
			this.txtVivienda.TabIndex = 48;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(16, 24);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(46, 16);
			this.label46.TabIndex = 47;
			this.label46.Text = "Vivienda";
			// 
			// ultraTabPageControl12
			// 
			this.ultraTabPageControl12.Controls.Add(this.ultraGrid6);
			this.ultraTabPageControl12.Controls.Add(this.ultraGrid5);
			this.ultraTabPageControl12.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl12.Name = "ultraTabPageControl12";
			this.ultraTabPageControl12.Size = new System.Drawing.Size(724, 335);
			// 
			// ultraGrid6
			// 
			this.ultraGrid6.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid6.DataMember = "RolObservacion";
			this.ultraGrid6.DataSource = this.cdsProvincia;
			appearance57.BackColor = System.Drawing.Color.White;
			this.ultraGrid6.DisplayLayout.Appearance = appearance57;
			this.ultraGrid6.DisplayLayout.AutoFitColumns = true;
			ultraGridBand22.AddButtonCaption = "PersFamiliar";
			ultraGridColumn167.Header.VisiblePosition = 1;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn167.Width = 195;
			ultraGridColumn168.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn168.Header.VisiblePosition = 2;
			ultraGridColumn168.Width = 507;
			ultraGridColumn169.Header.VisiblePosition = 0;
			ultraGridColumn169.Width = 168;
			ultraGridColumn170.Header.VisiblePosition = 3;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn170.Width = 181;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn167,
																											ultraGridColumn168,
																											ultraGridColumn169,
																											ultraGridColumn170});
			this.ultraGrid6.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.ultraGrid6.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid6.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid6.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance58.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid6.DisplayLayout.Override.CardAreaAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance59.FontData.BoldAsString = "True";
			appearance59.FontData.Name = "Arial";
			appearance59.FontData.SizeInPoints = 10F;
			appearance59.ForeColor = System.Drawing.Color.White;
			appearance59.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid6.DisplayLayout.Override.HeaderAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid6.DisplayLayout.Override.RowSelectorAppearance = appearance60;
			appearance61.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance61.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance61.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid6.DisplayLayout.Override.SelectedRowAppearance = appearance61;
			this.ultraGrid6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid6.Location = new System.Drawing.Point(16, 146);
			this.ultraGrid6.Name = "ultraGrid6";
			this.ultraGrid6.Size = new System.Drawing.Size(696, 130);
			this.ultraGrid6.TabIndex = 3;
			this.ultraGrid6.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid6_AfterRowInsert);
			// 
			// ultraGrid5
			// 
			this.ultraGrid5.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid5.DataMember = "RolUniforme";
			this.ultraGrid5.DataSource = this.cdsProvincia;
			appearance62.BackColor = System.Drawing.Color.White;
			this.ultraGrid5.DisplayLayout.Appearance = appearance62;
			this.ultraGrid5.DisplayLayout.AutoFitColumns = true;
			ultraGridBand23.AddButtonCaption = "PersFamiliar";
			ultraGridColumn171.Header.VisiblePosition = 0;
			ultraGridColumn171.Hidden = true;
			ultraGridColumn171.Width = 122;
			ultraGridColumn172.Header.VisiblePosition = 3;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn172.Width = 63;
			ultraGridColumn173.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn173.Header.VisiblePosition = 2;
			ultraGridColumn173.Width = 507;
			ultraGridColumn174.Format = "dd/MMM/yyyy";
			ultraGridColumn174.Header.VisiblePosition = 1;
			ultraGridColumn174.Width = 166;
			ultraGridColumn175.Header.VisiblePosition = 4;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn175.Width = 135;
			ultraGridBand23.Columns.AddRange(new object[] {
																											ultraGridColumn171,
																											ultraGridColumn172,
																											ultraGridColumn173,
																											ultraGridColumn174,
																											ultraGridColumn175});
			this.ultraGrid5.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.ultraGrid5.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid5.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid5.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance63.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid5.DisplayLayout.Override.CardAreaAppearance = appearance63;
			appearance64.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance64.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance64.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance64.FontData.BoldAsString = "True";
			appearance64.FontData.Name = "Arial";
			appearance64.FontData.SizeInPoints = 10F;
			appearance64.ForeColor = System.Drawing.Color.White;
			appearance64.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid5.DisplayLayout.Override.HeaderAppearance = appearance64;
			appearance65.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance65.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid5.DisplayLayout.Override.RowSelectorAppearance = appearance65;
			appearance66.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance66.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid5.DisplayLayout.Override.SelectedRowAppearance = appearance66;
			this.ultraGrid5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid5.Location = new System.Drawing.Point(18, 9);
			this.ultraGrid5.Name = "ultraGrid5";
			this.ultraGrid5.Size = new System.Drawing.Size(694, 130);
			this.ultraGrid5.TabIndex = 2;
			this.ultraGrid5.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid5_AfterRowInsert);
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
			this.cdsRolConcepto.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolConcepto_BeforeFill);
			// 
			// cdvSucursal
			// 
			this.cdvSucursal.BindingContextCtrl = this;
			this.cdvSucursal.DataSet = this.cdsSucursal;
			this.cdvSucursal.Sort = "Sucursal";
			this.cdvSucursal.TableName = "";
			this.cdvSucursal.TableViewName = "RolSucursal";
			// 
			// ultraTabControl1
			// 
			appearance67.BackColor = System.Drawing.SystemColors.ActiveBorder;
			this.ultraTabControl1.ClientAreaAppearance = appearance67;
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
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl11);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl12);
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 9);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(728, 361);
			this.ultraTabControl1.TabIndex = 0;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Personales";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Ubicacion";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Fechas";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Trabajo";
			ultraTab5.TabPage = this.ultraTabPageControl5;
			ultraTab5.Text = "Complemento";
			ultraTab6.TabPage = this.ultraTabPageControl6;
			ultraTab6.Text = "Adicional";
			ultraTab7.TabPage = this.ultraTabPageControl7;
			ultraTab7.Text = "Estudios";
			ultraTab8.TabPage = this.ultraTabPageControl8;
			ultraTab8.Text = "Familiares";
			ultraTab9.TabPage = this.ultraTabPageControl9;
			ultraTab9.Text = "Vacacion";
			ultraTab10.TabPage = this.ultraTabPageControl10;
			ultraTab10.Text = "Sueldo";
			ultraTab11.TabPage = this.ultraTabPageControl11;
			ultraTab11.Text = "Gastos";
			ultraTab12.TabPage = this.ultraTabPageControl12;
			ultraTab12.Text = "Uniformes";
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
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(724, 335);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Personal";
			this.barraDato1.DataNombreId = "idPersonal";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsPersonal;
			this.barraDato1.DataTabla = "Personal";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 416);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 2;
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
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtIdPersonal
			// 
			this.txtIdPersonal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idPersonal"));
			this.txtIdPersonal.Location = new System.Drawing.Point(488, 416);
			this.txtIdPersonal.Name = "txtIdPersonal";
			this.txtIdPersonal.PromptChar = ' ';
			this.txtIdPersonal.Size = new System.Drawing.Size(80, 22);
			this.txtIdPersonal.TabIndex = 3;
			// 
			// btNivel
			// 
			this.btNivel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNivel.Location = new System.Drawing.Point(320, 416);
			this.btNivel.Name = "btNivel";
			this.btNivel.Size = new System.Drawing.Size(24, 24);
			this.btNivel.TabIndex = 4;
			this.btNivel.Text = "N";
			this.toolTip1.SetToolTip(this.btNivel, "Niveles de Estudios");
			this.btNivel.Click += new System.EventHandler(this.btNivel_Click);
			// 
			// btTitulo
			// 
			this.btTitulo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTitulo.Location = new System.Drawing.Point(344, 416);
			this.btTitulo.Name = "btTitulo";
			this.btTitulo.Size = new System.Drawing.Size(24, 24);
			this.btTitulo.TabIndex = 5;
			this.btTitulo.Text = "T";
			this.toolTip1.SetToolTip(this.btTitulo, "Titulos de Estudios");
			this.btTitulo.Click += new System.EventHandler(this.btTitulo_Click);
			// 
			// btParentesco
			// 
			this.btParentesco.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btParentesco.Location = new System.Drawing.Point(368, 416);
			this.btParentesco.Name = "btParentesco";
			this.btParentesco.Size = new System.Drawing.Size(24, 24);
			this.btParentesco.TabIndex = 6;
			this.btParentesco.Text = "P";
			this.toolTip1.SetToolTip(this.btParentesco, "Parentescos Familiares");
			this.btParentesco.Click += new System.EventHandler(this.btParentesco_Click);
			// 
			// btCargasF
			// 
			this.btCargasF.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCargasF.Location = new System.Drawing.Point(392, 416);
			this.btCargasF.Name = "btCargasF";
			this.btCargasF.Size = new System.Drawing.Size(24, 24);
			this.btCargasF.TabIndex = 7;
			this.btCargasF.Text = "C";
			this.toolTip1.SetToolTip(this.btCargasF, "Genera Cargas Familiares");
			this.btCargasF.Click += new System.EventHandler(this.btCargasF_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Location = new System.Drawing.Point(416, 416);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 24);
			this.ultraButton1.TabIndex = 8;
			this.ultraButton1.Text = "R";
			this.toolTip1.SetToolTip(this.ultraButton1, "Reporte de Contratos");
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click_1);
			// 
			// btAuditoria
			// 
			appearance68.Image = ((object)(resources.GetObject("appearance68.Image")));
			this.btAuditoria.Appearance = appearance68;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(440, 416);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 24);
			this.btAuditoria.TabIndex = 97;
			this.toolTip1.SetToolTip(this.btAuditoria, "Auditoria de Registro");
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// btAdicionales
			// 
			this.btAdicionales.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAdicionales.Location = new System.Drawing.Point(712, 424);
			this.btAdicionales.Name = "btAdicionales";
			this.btAdicionales.Size = new System.Drawing.Size(24, 24);
			this.btAdicionales.TabIndex = 98;
			this.btAdicionales.Text = "O";
			this.toolTip1.SetToolTip(this.btAdicionales, "Organigrama");
			this.btAdicionales.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// lblPersonalDiscapacitado
			// 
			this.lblPersonalDiscapacitado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblPersonalDiscapacitado.AutoSize = true;
			this.lblPersonalDiscapacitado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPersonalDiscapacitado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPersonalDiscapacitado.Location = new System.Drawing.Point(16, 376);
			this.lblPersonalDiscapacitado.Name = "lblPersonalDiscapacitado";
			this.lblPersonalDiscapacitado.Size = new System.Drawing.Size(0, 26);
			this.lblPersonalDiscapacitado.TabIndex = 241;
			// 
			// RolPersonal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(752, 454);
			this.Controls.Add(this.lblPersonalDiscapacitado);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btCargasF);
			this.Controls.Add(this.btParentesco);
			this.Controls.Add(this.btTitulo);
			this.Controls.Add(this.btNivel);
			this.Controls.Add(this.txtIdPersonal);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.btAdicionales);
			this.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsPersonal, "Personal.Nombre"));
			this.MaximizeBox = false;
			this.Name = "RolPersonal";
			this.Text = "Datos del Personal";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolPersonal_Closing);
			this.Load += new System.EventHandler(this.RolPersonal_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCargas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoSectorial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optSistema)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAfiliacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optSexo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdReingreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolReingreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRegSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRegIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRelacionT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtNovedad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngreso)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbJefes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtJornal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSueldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalendario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDistrib)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsientoDistrib)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraOptionSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNumConadis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoDiscapacidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeDiscapacidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ultraTabPageControl8.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbHastaV)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesdeV)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid4)).EndInit();
			this.ultraTabPageControl11.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGasto)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor17)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtVestimenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEducacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlimentacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVivienda)).EndInit();
			this.ultraTabPageControl12.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIdPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void HabilitaTab()
		{
			//foreach(Infragistics.Win.UltraWinTabControl.UltraTab ctr in this.ultraTabControl1.Controls)
			foreach(Control ctr in this.ultraTabControl1.Controls)
				foreach(Control cn in ctr.Controls)
					this.barraDato1.HabilitaControl(cn, this.barraDato1.bEditar);
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private bool bFiltroSucursal = false;
		private void RolPersonal_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("79",cdsSeguridad))
			{
				MessageBox.Show("No puede ingresar a Personal", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			//validacion para reingresos 
			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'RPREING'")) 				
			{
				btnReingreso.Visible = true;
				grdReingreso.Visible = true;
			}

			cmbDesdeV.Value = DateTime.Today;
			cmbHastaV.Value = DateTime.Today;
			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Personal'";
			Funcion.EjecutaSQL(cdsAsientoDistrib, stAudita, true);
			Cursor = Cursors.Default;
			miAcceso =new Acceso("0501", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);

			txtIdPersonal.Width = 0;
			this.barraDato1.HabilitaControles(false);

			#region Filtro de Sucursales
			bFiltroSucursal = Funcion.bEjecutaSQL(cdsAsientoDistrib, 
				"Select IsNull(FiltroSucursal, 0) From SeteoRol");
			if (bFiltroSucursal)
			{
				string iSuc = Funcion.sEscalarSQL(cdsPersonal, "Exec VerificaUsuarioRol");
				if (iSuc != "")
				{
					barraDato1.stFiltroExt = " And  Personal.idRolSucursal in ( " + iSuc.ToString()+ " )";
					cdvSucursal.RowFilter = "idRolSucursal in (" + iSuc.ToString()+ " )";
				}
			}

			#endregion Filtro de Sucursales

			if (IdPersonal > 0)
			{
				BuscarPersona(IdPersonal);
			}
			else
			{
				this.barraDato1.PosUltima();
			}
			this.HabilitaTab();

			int iCuentaCaducados = Funcion.iEscalarSQL(cdsPersonal, "Exec RolCaducaCuenta");
			#region RolCaduca Contrato
			if (iCuentaCaducados > 0) 
			{ 
				MessageBox.Show("No tiene acceso a Eliminar Personal de Roles", "Información");
				RolContatoCaduca miCaducados = new RolContatoCaduca();
				miCaducados.ShowDialog();
			}
			#endregion RolCaduca Contrato
			#region Supervisor
			string stExec = "Select idPersonal,Nombre, Codigo from Personal where bSupervisor = 1";
			cmbSupervisor.DataSource = Funcion.dvProcedimiento(cdsPersonal, stExec);
			cmbSupervisor.DisplayLayout.Bands[0].Columns["idPersonal"].Hidden = true;
			#endregion Supervisor

			llenaDependencia();			
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Personal de Roles", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stSelect = "Select Count(*) From RolPago "
				+ "Where idPersonal = " + txtIdPersonal.Value.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsPersonal, stSelect, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede borrar personal porque existen "
					+ iCuenta.ToString() + " registros en rol de pagos", "Información");
				return;
			}
			stSelect = "Select Count(*) From RolPrestamo "
				+ "Where idPersonal = " + txtIdPersonal.Value.ToString();
			iCuenta = Funcion.iEscalarSQL(cdsPersonal, stSelect, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede borrar personal porque tiene "
					+ iCuenta.ToString() + " prestamos", "Información");
				return;
			}

			if (MessageBox.Show("¿Desea Borrar datos de ésta persona?",
				"Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
				== DialogResult.No) return;
			this.barraDato1.BorraRegistro();
		}

		private void BuscarPersona(int idPersona)
		{
			if (idPersona > 0) 
			{
				this.barraDato1.IdRegistro = idPersona;
				this.barraDato1.ProximoId(5);
			}
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
//			if (!miAcceso.Buscar)
//			{
//				MessageBox.Show("No tiene acceso a Buscar personal", "Información",
//					MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}

			using (BuscaPersonal miBusca = new BuscaPersonal())
			{
				if (DialogResult.OK == miBusca.ShowDialog())
				{
					try
					{
						int id = int.Parse(miBusca.idBusca.Text);
						BuscarPersona(id);
					}
					catch{}
				}
			}
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.dtIngreso, "");
			//this.errorProvider.SetError(this.txtApellido, "");
			this.errorProvider.SetError(this.txtNombre, "");
			this.errorProvider.SetError(this.txtCedula, "");
			this.errorProvider.SetError(this.cmbCargo, "");
			this.errorProvider.SetError(this.cmbDepartamento, "");
			this.errorProvider.SetError(this.txtRegIngreso, "");
			this.errorProvider.SetError(this.txtRegSalida, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.BorraErrores();
			bool bDato = false;
			bool bTrabajo = false;
			bool bFecha = false;
			if (dtIngreso.Value == DBNull.Value )
			{
				errorProvider.SetError(dtIngreso, "Ingrese la Fecha de Ingreso");
				bOk = false;
				bFecha = true;
			}
			else
			{
				DateTime dtFIngreso = (DateTime) dtIngreso.Value;
				if (dtFIngreso.Year < 1960)
				{
					errorProvider.SetError(dtIngreso, "Fecha de Ingreso no puede ser menor a 1960");
					bOk = false;
					bFecha = true;
				}
			}

			// Registro de Ingreso
//			if (txtRegIngreso.Text.Trim().Length == 0)
//			{
//				errorProvider.SetError(txtRegIngreso, "Ingrese Registro de Ingreso");
//				bOk = false;
//				bFecha = true;
//			}

			if (dtSalida.Value != DBNull.Value)
			{
				DateTime dtFSalida = (DateTime) dtSalida.Value;
				if (dtFSalida.Year < 2000)
				{
					errorProvider.SetError(dtSalida, "Fecha de Salida no puede ser menor a 2000");
					bOk = false;
					bFecha = true;
				}
				// Registro de Salida
				if (txtRegSalida.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtRegSalida, "Ingrese Registro de Salida");
					bOk = false;
					bFecha = true;
				}
			}

			if (txtNombre.Value == null || txtNombre.Value.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNombre, "Ingrese Nombre");
				bOk = false;
				bDato = true;
			}
			if (txtCedula.Value == null || txtCedula.Value.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtCedula, "Ingrese Cédula");
				bOk = false;
				bDato = true;
			}
			if (cmbCargo.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(cmbCargo, "Ingrese Cargo");
				bOk = false;
				bTrabajo = true;
			}
			if (cmbDepartamento.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(cmbDepartamento, "Ingrese Departamento");
				bOk = false;
				bTrabajo = true;
			}
			string stMensaje = "";
			if (bDato) stMensaje += "Error en Datos Personales. ";
			if (bFecha) stMensaje += "Error en Ficha de Fechas. ";
			if (bTrabajo) stMensaje += "Error en Datos de Trabajo. ";
			if (stMensaje.Trim().Length > 0)
				MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

			return bOk;
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a crear nuevo Personal", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			
			this.barraDato1.CrearRegistro();
			this.HabilitaTab();
			cmbCedula.Enabled = true;
//			this.ultraTabPageControl1.Focus();
			txtNombre.SelectAll();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.BorraErrores();
			this.barraDato1.DeshacerRegistro();
			this.HabilitaTab();
			cmbCedula.Enabled = false;
			bytFoto = null;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a editar Personal", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.barraDato1.EditarRegistro();
			this.HabilitaTab();
			cmbCedula.Enabled = true;
			bytFoto = null;
		  cmbSupervisor.Enabled = !this.chkSupervisor.Checked;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			//Pone en Codigo de Empleado Un codigo que toma de SeteoG
			C1.Data.C1DataRow drPer;
			drPer = this.cdsPersonal.TableViews["Personal"].Rows[0];
			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'RPCODF'")) 				
			{		
				string stSelect ="select Valor from seteog where codigo like 'RPCODF'";
				double stValor = Funcion.dEscalarSQL(cdsPersonal,stSelect);
						drPer["RegEntrada"] = stValor.ToString();
			}
			//Validacion de cedula al grabar segun codigo configuracion
//			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'RPVCI'")) 				
//			{		
				if (cmbCedula.Value.ToString()=="2")
					if(Funcion.VerificaCedula(this.txtCedula.Text, true).Length > 0)
					{
						string stMensaje = Funcion.VerificaCedula(this.txtCedula.Text.Trim(), true);
						MessageBox.Show(stMensaje, "Informacion",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
						return;
					}
//			}

			if (!Verifica()) return;

			string sSQL = string.Format("Exec ValidaExistenciaEmpleados '{0}', '{1}', '{2}', {3}", this.txtCedula.Text.ToString(), this.txtCodigo.Text.ToString(), this.txtNombre.Text.ToString(), (int)drPer["idPersonal"]);
			string sMensaje = Funcion.sEscalarSQL(cdsPersonal, sSQL);
			if (!sMensaje.StartsWith("CORRECTO"))
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ultraTabControl1.Tabs[0].Selected = true;				
				return;
			}
	
			if ((double)this.txtSueldo.Value == 0)
			{
				MessageBox.Show("Ingrese el Sueldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.ultraTabControl1.Tabs[3].Selected = true;
				this.txtSueldo.Focus();
				return;
			}

			if (this.ultraTextEditor4.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Estado Civil", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.ultraTabControl1.Tabs[4].Selected = true;
				this.ultraTextEditor4.Focus();
				return;
			}

			if (this.dtNacimiento.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha de Nacimiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.ultraTabControl1.Tabs[2].Selected = true;
				this.dtNacimiento.Focus();
				return;
			}

			int iCuenta = -1;
			if (barraDato1.bNuevo)
			{
				string stSelect = "Select Count(*) From Personal";
				iCuenta = Funcion.iEscalarSQL(cdsPersonal, stSelect, true);
			}

			this.barraDato1.GrabaRegistro();

			if (iCuenta >= 0)
			{
				int IdPersonal = (int) txtIdPersonal.Value;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (dr.Cells["idPersonal"].Value != System.DBNull.Value) dr.Cells["idPersonal"].Value = IdPersonal;
				}
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				{
					if (dr.Cells["idPersonal"].Value != System.DBNull.Value) dr.Cells["idPersonal"].Value = IdPersonal;
				}
			}

			cdsProvincia.Update();
			cdsRolReingreso.Update();

			this.HabilitaTab();
			cmbCedula.Enabled = false;
//			GrabaFoto();
//			bytFoto = null;
		}

		private void Imprimir(int iTipo)
		{
			Cursor = Cursors.WaitCursor;
			
			#region Imprime el reporte
			ParameterFields paramFields = new ParameterFields ();

			Reporte Reporte = new Reporte("NominaDePersonal.rpt", "");
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Titulo("Nomina");
				// Muestra en pantalla
			if (iTipo == 0) Reporte.ShowDialog();
			if (iTipo >= 1) // Imprime una o mas copias
			{
				string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
				Reporte.ImprimeRep(1);					
			}			
			#endregion Imprime el reporte
		
			Cursor = Cursors.Default;
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			frmReportePersonal Reporte = new frmReportePersonal();
			Reporte.MdiParent = this.MdiParent;
			Reporte.Show();

			//Imprimir(0);

			//			RolPersonalImprime miRep = new RolPersonalImprime();
			//			miRep.MdiParent = this.MdiParent;
			//			miRep.Show();

			//			Cursor = Cursors.WaitCursor;
			//			Reporte miRep = new Reporte("Personal");
			//			miRep.MdiParent = this.MdiParent;
			//			miRep.Show();
			//			Cursor = Cursors.Default;
		}

		private void ActualizaTablas()
		{
			this.cdsSucursal.Fill(true);
		}

		private void btSucursal_Click(object sender, System.EventArgs e)
		{
			RolSucursal miSucursal = new RolSucursal();
			miSucursal.MdiParent = this.MdiParent;
			miSucursal.Show();
			this.ActualizaTablas();
		}

		private void btDepartamento_Click(object sender, System.EventArgs e)
		{
			RolDepartamento miDepartamento = new RolDepartamento();
			miDepartamento.MdiParent = this.MdiParent;
			miDepartamento.Show();	
			this.ActualizaTablas();
		}

		private void btCargo_Click(object sender, System.EventArgs e)
		{
			RolCargo miCargo = new RolCargo();
			miCargo.MdiParent = this.MdiParent;
			miCargo.Show();
			this.ActualizaTablas();
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsSucursal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSucursal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void txtCedula_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				string stMensaje = Funcion.VerificaCedula(this.txtCedula.Text.Trim(), true);
				MessageBox.Show(stMensaje, "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		private void btGrupo_Click(object sender, System.EventArgs e)
		{
			RolGrupo miGrupo = new RolGrupo();
			miGrupo.MdiParent = this.MdiParent;
			miGrupo.Show();
		}

		private void cdsCiudad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCiudad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsRolConcepto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRolConcepto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsProvincia_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProvincia.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			int IdPersonal = -1;
			if (txtIdPersonal.Value != System.DBNull.Value)
				IdPersonal = (int) txtIdPersonal.Value;
			C1.Data.FilterCondition fcEstudios = new 
				C1.Data.FilterCondition(cdsProvincia, "PersEstudios", "idPersonal = " + IdPersonal.ToString());
			C1.Data.FilterCondition fcFamiliar = new 
				C1.Data.FilterCondition(cdsProvincia, "PersFamiliar", "idPersonal = " + IdPersonal.ToString());
			C1.Data.FilterCondition fcSueldo = new 
				C1.Data.FilterCondition(cdsProvincia, "RolSueldo", "idPersonal = " + IdPersonal.ToString());
			C1.Data.FilterCondition fcVacacion = new 
				C1.Data.FilterCondition(cdsProvincia, "RolVacacion", "idPersonal = " + IdPersonal.ToString());
			C1.Data.FilterCondition fcUniforme = new 
				C1.Data.FilterCondition(cdsProvincia, "RolUniforme", "idPersonal = " + IdPersonal.ToString());
			C1.Data.FilterCondition fcObservacion = new 
				C1.Data.FilterCondition(cdsProvincia, "RolObservacion", "idPersonal = " + IdPersonal.ToString());
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();

			fcTotal.Add(fcSueldo);
			fcTotal.Add(fcVacacion);
			fcTotal.Add(fcEstudios);
			fcTotal.Add(fcFamiliar);
			fcTotal.Add(fcUniforme);
			fcTotal.Add(fcObservacion);

			cdsProvincia.Fill(fcTotal, false);

			C1.Data.FilterCondition fcRolreingreso = new 
				C1.Data.FilterCondition(cdsRolReingreso, "RolReingreso", "idPersonal = " + IdPersonal.ToString());
			cdsRolReingreso.Fill(fcRolreingreso, false);
			llenaDependencia();

			this.lblPersonalDiscapacitado.Text = Funcion.sEscalarSQL(cdsPersonal, "Exec NumeroPersonalDiscapacitado");		
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void btNivel_Click(object sender, System.EventArgs e)
		{
			RolNivel miNivel = new RolNivel();
			miNivel.ShowDialog();
		}

		private void btTitulo_Click(object sender, System.EventArgs e)
		{
			RolTitulo miTitulo = new RolTitulo();
			miTitulo.ShowDialog();		
		}

		private void btParentesco_Click(object sender, System.EventArgs e)
		{
			RolParentesco miParentesco = new RolParentesco();
			miParentesco.ShowDialog();		
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPersonal"].Value = (int) this.txtIdPersonal.Value;
		}

		private void ultraGrid2_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPersonal"].Value = (int) this.txtIdPersonal.Value;		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
		
		}
		private byte[] bytFoto;
		public void LeeFoto()
		{
			string stArchivoFoto = "";
			if (this.openFileDialog.ShowDialog() == DialogResult.OK)
			{
				stArchivoFoto = this.openFileDialog.FileName;
			}
			else 
			{
				bytFoto = null;
				return;
			}

			FileStream fs = new FileStream(stArchivoFoto, FileMode.Open, FileAccess.Read);
			BinaryReader br = new BinaryReader(fs);

			bytFoto = br.ReadBytes((int)fs.Length);

			br.Close();
			fs.Close();
		}

		private void GrabaFoto()
		{
			if (bytFoto == null) return;
			string stSelect = "UPDATE Personal SET Fotos = @Photo Where idPersonal = " 
				+ txtIdPersonal.Value.ToString();
			C1.Data.SchemaObjects.Connection miConeccion = cdsPersonal.Schema.Connections[0];
			//			miConeccion.CommandTimeout = 300;
			miConeccion.Open();
			SqlCommand cmdGrabaFoto = new SqlCommand (stSelect,	
				(SqlConnection) miConeccion.DbConnection);
			cmdGrabaFoto.Parameters.Add("@Photo", SqlDbType.Image, bytFoto.Length).Value = bytFoto;

			try
			{
				cmdGrabaFoto.ExecuteNonQuery();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			miConeccion.Close();

			barraDato1.ProximoId(5);
		}
	  private void picFoto_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			LeeFoto();		
		}

		private void RolPersonal_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Personal'";
			Funcion.EjecutaSQL(cdsAsientoDistrib, stAudita, true);
		}

		private void btCargasF_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea Recalcular el numero de cargas en base a familiares", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stExec = "Exec RolCargas";
			Funcion.EjecutaSQL(cdsPersonal, stExec, true);
			MessageBox.Show("Numero de cargas Recalculado", "Confirmación",
				MessageBoxButtons.OK, MessageBoxIcon.Question);
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


		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid2.ActiveCell == null) return;
			try
			{
				if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid2.ActiveCell.DroppedDown) return;

						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid2.ActiveCell.DroppedDown) return;
			
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
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

		private void btHorario_Click(object sender, System.EventArgs e)
		{
			CambiaHorario ch = new CambiaHorario();
			ch.MdiParent = this.MdiParent;
			ch.Show();
		}

		private void btHorario1_Click(object sender, System.EventArgs e)
		{
			int IdRolHorario = 0;
			if (cmbCalendario.Text.Trim().Length > 0)
			{
				IdRolHorario = (int) cmbCalendario.Value;
				RolHorario miHorario = new RolHorario(IdRolHorario);
				miHorario.MdiParent = this.MdiParent;
				miHorario.Show();
			}
			else
			{
				RolHorario miHorario = new RolHorario();
				miHorario.MdiParent = this.MdiParent;
				miHorario.Show();
			}
		}

		private void btHoraExtra_Click(object sender, System.EventArgs e)
		{
			int IdPersonal = (int) txtIdPersonal.Value;
			RolHoraExtra miHorario = new RolHoraExtra(IdPersonal);
			miHorario.MdiParent = this.MdiParent;
			miHorario.Show();	
		}

		private void ultraButton1_Click_1(object sender, System.EventArgs e)
		{
			RolContrato miContrato = new RolContrato();
			miContrato.MdiParent = this.MdiParent;
			miContrato.Show();
		}

		private void btCiudad_Click(object sender, System.EventArgs e)
		{
			ClienteCiudad miCiudad = new ClienteCiudad();
			miCiudad.MdiParent = this.MdiParent;
			miCiudad.Show();
		}

		private void btProvincia_Click(object sender, System.EventArgs e)
		{
			RolProvincia miProvincia = new RolProvincia();
			miProvincia.MdiParent = this.MdiParent;
			miProvincia.Show();
		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(35, int.Parse(txtIdPersonal.Text.ToString()), true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(35, int.Parse(txtIdPersonal.Value.ToString()), false);
			Cursor = Cursors.Default;
		}

		private void ultraTabPageControl5_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void ultraGrid3_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPersonal"].Value = (int) this.txtIdPersonal.Value;		
		}

		private void ultraGrid4_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPersonal"].Value = (int) this.txtIdPersonal.Value;		
		}

		private void btTotalGasto_Click(object sender, System.EventArgs e)
		{
			double dTotal = (double) txtVestimenta.Value + (double) txtAlimentacion.Value + 
				(double) txtSalud.Value + (double) txtVivienda.Value + (double) txtEducacion.Value;
			txtTotalGasto.Value = Math.Round(dTotal / 12.0, 2);
		}

		private void btRepVacacion_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesdeV.Value;
			DateTime dtHasta = (DateTime) cmbHastaV.Value;
			string stFiltro = "{RolVacacion.idPersonal} = "
				+ txtIdPersonal.Value.ToString() + " And {RolVacacion.Fecha} >= #"
				+ dtDesde.ToString("MM/dd/yyyy") + "# And {RolVacacion.Fecha} < #"
				+ dtHasta.AddDays(1).ToString("MM/dd/yyyy") + "#";
			Reporte miRep = new Reporte("RolVacacion.rpt", stFiltro);
			miRep.MdiParent = this.MdiParent;
			Cursor = Cursors.Default;
			miRep.Show();
		}

		private void btRepVacacion_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;

			DateTime dtDesde = (DateTime) cmbDesdeV.Value;
			DateTime dtHasta = (DateTime) cmbHastaV.Value;

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@idPersonal";
			discreteVal.Value = int.Parse(txtIdPersonal.Value.ToString());
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramFieldD = new ParameterField ();
			ParameterDiscreteValue discreteValD = new ParameterDiscreteValue ();
			paramFieldD.ParameterFieldName = "@FechaDesde";
			discreteValD.Value = dtDesde.ToString("yyyy-MM-dd");
			paramFieldD.CurrentValues.Add (discreteValD);
			paramFields.Add (paramFieldD);
			ParameterField paramFieldH = new ParameterField ();
			ParameterDiscreteValue discreteValH = new ParameterDiscreteValue ();
			paramFieldH.ParameterFieldName = "@FechaHasta";
			discreteValH.Value = dtHasta.ToString("yyyy-MM-dd");
			paramFieldH.CurrentValues.Add (discreteValH);
			paramFields.Add (paramFieldH);
			Reporte miRep = new Reporte("RolVacacionSP.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			Cursor = Cursors.Default;
		}
		private void ultraGrid5_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPersonal"].Value = (int) this.txtIdPersonal.Value;
		}
		private void label63_Click(object sender, System.EventArgs e)
		{
		}
		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			RolTipoContrato miTipoContrato = new RolTipoContrato();
			miTipoContrato.MdiParent = this.MdiParent;
			miTipoContrato.Show();
		}

		private void cdsTipoContrato_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoContrato.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void ultraGrid6_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPersonal"].Value = (int) this.txtIdPersonal.Value;
		}

		private void ultraTabPageControl2_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void btnReingreso_Click(object sender, System.EventArgs e)
		{			
			if (barraDato1.bEditar == false)
			{
				MessageBox.Show("No puede generar Reingresos sin Editar", "Información");
				return;
			}
			C1.Data.C1DataRow drPer1;
			drPer1 = this.cdsPersonal.TableViews["Personal"].Rows[0];	
	
			dtSalida.Value = drPer1["FechaSalida"];


			if (dtSalida.Value == DBNull.Value) 
			{
				MessageBox.Show("Fecha de salida Vacia","NO puede Generar Reingreso",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
				return;
			}
			string stSelect = " Exec ProcedRolReingreso "+  this.txtIdPersonal.Value + ", 1";
			try
			{
				Funcion.EjecutaSQL(cdsPersonal,stSelect);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Error al Generar...\n"+ex);
			}
			      				
			drPer1["FechaSalida"] = DBNull.Value;

		}

		private void cdsRolReingreso_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolReingreso.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void ultraTabPageControl3_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void chkSupervisor_CheckedChanged(object sender, System.EventArgs e)
		{
		this.cmbSupervisor.Enabled = !this.chkSupervisor.Checked;
			if(chkSupervisor.Checked == true)
			{
			cmbSupervisor.Value = 0;	
			}

		}

		private void btDistribucion_Click(object sender, System.EventArgs e)
		{
			using (AsientoDist miDist = new AsientoDist())
			{
				miDist.ShowDialog();
			}
			cdsAsientoDistrib.Clear();
			cdsAsientoDistrib.Fill();
		}

		private void ultraCombo3_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbDistrib.Value = 0;
		}

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			P_RolOrganigrama miRol = new P_RolOrganigrama();
			miRol.Show();
		}

		private void cmbJefes_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				C1.Data.C1DataRow drPer1;
				drPer1 = this.cdsPersonal.TableViews["Personal"].Rows[0];	

				if (cmbJefes.ActiveRow != null)					
				drPer1["idJefe"] = 0;
			}
		}
		private void llenaDependencia()
		{
			string stSelect =" Select idPersonal, p.Nombre,rc.Descripcion from Personal p "+
				" inner join RolCargo rc on rc.idCargo = p.idCargo";
			cmbJefes.DataSource = Funcion.dvProcedimiento(cdsPersonal,stSelect);
			cmbJefes.ValueMember = "idPersonal";
			cmbJefes.DisplayMember = "Nombre";
			cmbJefes.DisplayLayout.Bands[0].Columns["idPersonal"].Hidden = true;
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor5_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label38_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label36_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label35_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label31_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTextEditor8_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label30_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTextEditor7_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label29_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTextEditor6_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label28_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTextEditor5_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor4_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label37_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
