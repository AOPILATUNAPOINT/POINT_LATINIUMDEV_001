using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepCompInv.
	/// </summary>
	public class RepCompCredito : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btInventario;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.Misc.UltraButton btKardex;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tcGrillas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private C1.Data.C1DataSet cdsGrupo;
		private Infragistics.Win.Misc.UltraButton btFiltroGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFiltro;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbAndOr;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSector;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoCli;
		private Infragistics.Win.Misc.UltraButton btPersonal;
		private Infragistics.Win.Misc.UltraButton btSector;
		private Infragistics.Win.Misc.UltraButton btCiudad;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataView cdvReporte;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoCliente;
		private System.ComponentModel.IContainer components;

		public RepCompCredito()
		{
			InitializeComponent();
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
			this.components = new System.ComponentModel.Container();
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteSector", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteCiudad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.btPersonal = new Infragistics.Win.Misc.UltraButton();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsVendedor = new C1.Data.C1DataSet();
			this.cmbAndOr = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.btSector = new Infragistics.Win.Misc.UltraButton();
			this.cmbSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.btCiudad = new Infragistics.Win.Misc.UltraButton();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btFiltroGrupo = new Infragistics.Win.Misc.UltraButton();
			this.txtFiltro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbGrupoCli = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btInventario = new Infragistics.Win.Misc.UltraButton();
			this.cmbDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.tcGrillas = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btKardex = new Infragistics.Win.Misc.UltraButton();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.chkTodoCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAndOr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCli)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tcGrillas)).BeginInit();
			this.tcGrillas.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(572, 254);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(16, 16);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(544, 224);
			this.ultraGrid1.TabIndex = 0;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(572, 254);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Location = new System.Drawing.Point(16, 16);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(544, 224);
			this.ultraGrid2.TabIndex = 1;
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.chkTodoCliente);
			this.ultraTabPageControl4.Controls.Add(this.btReporte);
			this.ultraTabPageControl4.Controls.Add(this.cmbReporte);
			this.ultraTabPageControl4.Controls.Add(this.cmbTipo);
			this.ultraTabPageControl4.Controls.Add(this.btPersonal);
			this.ultraTabPageControl4.Controls.Add(this.cmbPersonal);
			this.ultraTabPageControl4.Controls.Add(this.cmbAndOr);
			this.ultraTabPageControl4.Controls.Add(this.btSector);
			this.ultraTabPageControl4.Controls.Add(this.cmbSector);
			this.ultraTabPageControl4.Controls.Add(this.btCiudad);
			this.ultraTabPageControl4.Controls.Add(this.cmbCiudad);
			this.ultraTabPageControl4.Controls.Add(this.btFiltroGrupo);
			this.ultraTabPageControl4.Controls.Add(this.txtFiltro);
			this.ultraTabPageControl4.Controls.Add(this.cmbGrupoCli);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(2, 24);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(572, 254);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(528, 24);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(24, 24);
			this.btReporte.TabIndex = 47;
			this.btReporte.Text = "...";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			this.cmbReporte.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 118;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 54;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 54;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(384, 24);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(136, 21);
			this.cmbReporte.TabIndex = 46;
			this.toolTip1.SetToolTip(this.cmbReporte, "Grupo 100 y Tipo Avanzado");
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 10 And idGrupo = 100";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// cmbTipo
			// 
			this.cmbTipo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "CxCobrar";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "CxPagar";
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Todo";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Items.Add(valueListItem2);
			this.cmbTipo.Items.Add(valueListItem3);
			this.cmbTipo.Location = new System.Drawing.Point(40, 24);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(96, 21);
			this.cmbTipo.TabIndex = 14;
			this.cmbTipo.Text = "CxCobrar";
			// 
			// btPersonal
			// 
			this.btPersonal.Location = new System.Drawing.Point(176, 152);
			this.btPersonal.Name = "btPersonal";
			this.btPersonal.Size = new System.Drawing.Size(32, 24);
			this.btPersonal.TabIndex = 13;
			this.btPersonal.Text = "->";
			this.btPersonal.Click += new System.EventHandler(this.btTipoGrupo_Click);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataMember = "Personal";
			this.cmbPersonal.DataSource = this.cdsVendedor;
			appearance11.BackColor = System.Drawing.Color.White;
			this.cmbPersonal.DisplayLayout.Appearance = appearance11;
			ultraGridBand2.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 8;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Width = 210;
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 11;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 12;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 15;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 16;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 17;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 18;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 19;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 20;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 21;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 22;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 23;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 24;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 25;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 26;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 27;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 28;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 29;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 30;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 31;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 32;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 33;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 34;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 35;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 36;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 37;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 38;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 39;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 40;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 41;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 42;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 43;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 44;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 45;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 46;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 47;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 48;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 49;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 50;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 51;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 52;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 53;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 54;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 55;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 56;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 57;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 58;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 59;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 60;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 61;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 62;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 63;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 64;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 65;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 66;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 67;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 68;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 69;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 70;
			ultraGridColumn79.Header.VisiblePosition = 71;
			ultraGridColumn80.Header.VisiblePosition = 72;
			ultraGridColumn81.Header.VisiblePosition = 73;
			ultraGridColumn82.Header.VisiblePosition = 74;
			ultraGridColumn83.Header.VisiblePosition = 75;
			ultraGridColumn84.Header.VisiblePosition = 76;
			ultraGridColumn85.Header.VisiblePosition = 77;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn85});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.cmbPersonal.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 10F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbPersonal.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbPersonal.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbPersonal.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPersonal.Location = new System.Drawing.Point(40, 152);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(136, 21);
			this.cmbPersonal.TabIndex = 12;
			this.cmbPersonal.ValueMember = "idPersonal";
			// 
			// cdsVendedor
			// 
			this.cdsVendedor.BindingContextCtrl = this;
			this.cdsVendedor.DataLibrary = "LibPersonal";
			this.cdsVendedor.DataLibraryUrl = "";
			this.cdsVendedor.DataSetDef = "dsPersonal";
			this.cdsVendedor.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsVendedor.Name = "cdsVendedor";
			this.cdsVendedor.SchemaClassName = "LibPersonal.DataClass";
			this.cdsVendedor.SchemaDef = null;
			// 
			// cmbAndOr
			// 
			this.cmbAndOr.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem4.DataValue = "And";
			valueListItem4.DisplayText = "And";
			valueListItem5.DataValue = "Or";
			valueListItem5.DisplayText = "Or";
			this.cmbAndOr.Items.Add(valueListItem4);
			this.cmbAndOr.Items.Add(valueListItem5);
			this.cmbAndOr.Location = new System.Drawing.Point(152, 24);
			this.cmbAndOr.Name = "cmbAndOr";
			this.cmbAndOr.Size = new System.Drawing.Size(48, 21);
			this.cmbAndOr.TabIndex = 9;
			this.cmbAndOr.Text = "And";
			// 
			// btSector
			// 
			this.btSector.Location = new System.Drawing.Point(176, 120);
			this.btSector.Name = "btSector";
			this.btSector.Size = new System.Drawing.Size(32, 24);
			this.btSector.TabIndex = 8;
			this.btSector.Text = "->";
			this.btSector.Click += new System.EventHandler(this.btMarca_Click);
			// 
			// cmbSector
			// 
			this.cmbSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSector.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSector.DataMember = "ClienteSector";
			this.cmbSector.DataSource = this.cdsGrupo;
			ultraGridBand3.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn86.Header.VisiblePosition = 0;
			ultraGridColumn87.Header.VisiblePosition = 1;
			ultraGridColumn88.Header.VisiblePosition = 2;
			ultraGridColumn88.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88});
			this.cmbSector.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSector.DisplayMember = "Sector";
			this.cmbSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSector.Location = new System.Drawing.Point(40, 120);
			this.cmbSector.Name = "cmbSector";
			this.cmbSector.Size = new System.Drawing.Size(136, 21);
			this.cmbSector.TabIndex = 7;
			this.cmbSector.ValueMember = "idSector";
			// 
			// cdsGrupo
			// 
			this.cdsGrupo.BindingContextCtrl = this;
			this.cdsGrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupo.DataLibraryUrl = "";
			this.cdsGrupo.DataSetDef = "dsClienteTabla";
			this.cdsGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupo.Name = "cdsGrupo";
			this.cdsGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupo.SchemaDef = null;
			this.cdsGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupo_BeforeFill);
			// 
			// btCiudad
			// 
			this.btCiudad.Location = new System.Drawing.Point(176, 88);
			this.btCiudad.Name = "btCiudad";
			this.btCiudad.Size = new System.Drawing.Size(32, 24);
			this.btCiudad.TabIndex = 6;
			this.btCiudad.Text = "->";
			this.btCiudad.Click += new System.EventHandler(this.btSubGrupo_Click);
			// 
			// cmbCiudad
			// 
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataMember = "ClienteCiudad";
			this.cmbCiudad.DataSource = this.cdsGrupo;
			appearance16.BackColor = System.Drawing.Color.White;
			this.cmbCiudad.DisplayLayout.Appearance = appearance16;
			ultraGridBand4.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn89.Header.VisiblePosition = 1;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 0;
			ultraGridColumn90.Width = 194;
			ultraGridColumn91.Header.VisiblePosition = 2;
			ultraGridColumn91.Width = 77;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.cmbCiudad.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 10F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbCiudad.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCiudad.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCiudad.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(40, 88);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(136, 21);
			this.cmbCiudad.TabIndex = 5;
			this.cmbCiudad.ValueMember = "idCiudad";
			// 
			// btFiltroGrupo
			// 
			this.btFiltroGrupo.Location = new System.Drawing.Point(176, 56);
			this.btFiltroGrupo.Name = "btFiltroGrupo";
			this.btFiltroGrupo.Size = new System.Drawing.Size(32, 24);
			this.btFiltroGrupo.TabIndex = 4;
			this.btFiltroGrupo.Text = "->";
			this.btFiltroGrupo.Click += new System.EventHandler(this.btFiltroGrupo_Click);
			// 
			// txtFiltro
			// 
			this.txtFiltro.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFiltro.Location = new System.Drawing.Point(232, 56);
			this.txtFiltro.Multiline = true;
			this.txtFiltro.Name = "txtFiltro";
			this.txtFiltro.Size = new System.Drawing.Size(320, 184);
			this.txtFiltro.TabIndex = 3;
			// 
			// cmbGrupoCli
			// 
			this.cmbGrupoCli.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoCli.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoCli.DataMember = "ClienteGrupo";
			this.cmbGrupoCli.DataSource = this.cdsGrupo;
			appearance21.BackColor = System.Drawing.Color.White;
			this.cmbGrupoCli.DisplayLayout.Appearance = appearance21;
			ultraGridBand5.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn92.Header.VisiblePosition = 0;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 1;
			ultraGridColumn93.Width = 206;
			ultraGridColumn94.Header.VisiblePosition = 2;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 3;
			ultraGridColumn95.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95});
			this.cmbGrupoCli.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.cmbGrupoCli.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 10F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbGrupoCli.DisplayLayout.Override.HeaderAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupoCli.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupoCli.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.cmbGrupoCli.DisplayMember = "Grupo";
			this.cmbGrupoCli.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoCli.Location = new System.Drawing.Point(40, 56);
			this.cmbGrupoCli.Name = "cmbGrupoCli";
			this.cmbGrupoCli.Size = new System.Drawing.Size(136, 21);
			this.cmbGrupoCli.TabIndex = 1;
			this.cmbGrupoCli.ValueMember = "idGrupoCliente";
			// 
			// btInventario
			// 
			this.btInventario.Location = new System.Drawing.Point(352, 16);
			this.btInventario.Name = "btInventario";
			this.btInventario.Size = new System.Drawing.Size(64, 24);
			this.btInventario.TabIndex = 0;
			this.btInventario.Text = "Cliente";
			this.btInventario.Click += new System.EventHandler(this.btInventario_Click);
			// 
			// cmbDesde
			// 
			this.cmbDesde.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(80, 16);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbDesde.TabIndex = 2;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(43, 14);
			this.ultraLabel1.TabIndex = 3;
			this.ultraLabel1.Text = "Periodo";
			// 
			// cmbHasta
			// 
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(200, 16);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(112, 21);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 4;
			// 
			// tcGrillas
			// 
			this.tcGrillas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tcGrillas.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tcGrillas.Controls.Add(this.ultraTabPageControl1);
			this.tcGrillas.Controls.Add(this.ultraTabPageControl2);
			this.tcGrillas.Controls.Add(this.ultraTabPageControl4);
			this.tcGrillas.Location = new System.Drawing.Point(24, 48);
			this.tcGrillas.Name = "tcGrillas";
			this.tcGrillas.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tcGrillas.Size = new System.Drawing.Size(576, 280);
			this.tcGrillas.TabIndex = 5;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Principal";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Detalle";
			ultraTab3.TabPage = this.ultraTabPageControl4;
			ultraTab3.Text = "Opciones";
			this.tcGrillas.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																												ultraTab1,
																																												ultraTab2,
																																												ultraTab3});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(572, 254);
			// 
			// btKardex
			// 
			this.btKardex.Location = new System.Drawing.Point(424, 16);
			this.btKardex.Name = "btKardex";
			this.btKardex.Size = new System.Drawing.Size(64, 24);
			this.btKardex.TabIndex = 6;
			this.btKardex.Text = "Factura";
			this.btKardex.Click += new System.EventHandler(this.btKardex_Click);
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(496, 16);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(64, 24);
			this.btFactura.TabIndex = 7;
			this.btFactura.Text = "Comprobante";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// chkTodoCliente
			// 
			this.chkTodoCliente.Location = new System.Drawing.Point(240, 24);
			this.chkTodoCliente.Name = "chkTodoCliente";
			this.chkTodoCliente.Size = new System.Drawing.Size(96, 24);
			this.chkTodoCliente.TabIndex = 48;
			this.chkTodoCliente.Text = "Todo Cliente";
			// 
			// RepCompCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(616, 342);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.btKardex);
			this.Controls.Add(this.tcGrillas);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.btInventario);
			this.Name = "RepCompCredito";
			this.Text = "Comparacion Estado de Cuenta";
			this.Load += new System.EventHandler(this.RepCompInv_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAndOr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoCli)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tcGrillas)).EndInit();
			this.tcGrillas.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSet = new DataSet();
		DataSet dsDetalle = new DataSet();

		private void btInventario_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stStore = "CreditoGenera";
			string stProcedimiento = "EXEC " + stStore + " '" + 
				dtDesde.ToString("yyyyMMdd") + "', '" + dtHasta.ToString("yyyyMMdd") + "', ";
			stProcedimiento += "'" + txtFiltro.Text.ToString().Trim() + "', 0, "
				+ cmbTipo.Value.ToString();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			tcGrillas.Tabs[0].Selected = true;
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void RepCompInv_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
		}

		private void btKardex_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null && !chkTodoCliente.Checked)
			{
				MessageBox.Show("Escoja Cliente" ,"Información");
				return;
			}
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stStore = "CreditoDetalleGenera";
			string stIdCliente = "0";
			if (ultraGrid1.ActiveRow != null && !chkTodoCliente.Checked)
				stIdCliente = ultraGrid1.ActiveRow.Cells["idCliente"].Value.ToString();
			string stCliente = "', " + stIdCliente;
			string stProcedimiento = "EXEC " + stStore + " '" + 
				dtDesde.ToString("yyyyMMdd") + "', '" + dtHasta.ToString("yyyyMMdd") + stCliente;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			dsDetalle.Reset();
			dsDetalle.Clear();
			try
			{
				myData.Fill(dsDetalle);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			tcGrillas.Tabs[1].Selected = true;
			ultraGrid2.DataSource = dsDetalle.Tables[0].DefaultView;
			ultraGrid2.DataBind();	
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid2.ActiveRow == null)
			{
				MessageBox.Show("Escoja Linea de pago" ,"Información");
				return;
			}
			int IdCompra = int.Parse(ultraGrid2.ActiveRow.Cells["idCompra"].Text);
			if (IdCompra == 0)
			{
				MessageBox.Show("Linea de pago no tiene una factura" ,"Información");
				return;
			}
			int IdTipoFactura = int.Parse(ultraGrid2.ActiveRow.Cells["idTipo"].Text);
			if (IdTipoFactura < 0) // Pago
			{
//				Pagos miPago = new Pagos(-IdTipoFactura, IdCompra);
//				miPago.MdiParent = this.MdiParent;
//				miPago.Show();
				return;
			}
//			Compra miCompra = new Compra(IdTipoFactura, IdCompra);
//			miCompra.MdiParent = this.MdiParent;
//			miCompra.Show();
		}

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btFiltroGrupo_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " cl.idGrupoCliente = " + cmbGrupoCli.Value.ToString();
		}

		private void btSubGrupo_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " cl.idCiudad = " + cmbCiudad.Value.ToString();
		}

		private void btMarca_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " cl.idSector = " + cmbSector.Value.ToString();
		}

		private void btTipoGrupo_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " cl.idVendedor = " + cmbPersonal.Value.ToString();
		}

		private void btArtSri_Click(object sender, System.EventArgs e)
		{
			txtFiltro.Text += "  " + cmbAndOr.Text + " cl.idVendedor = " + cmbTipo.Value.ToString();
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null && !chkTodoCliente.Checked)
			{
				MessageBox.Show("Escoja un cliente en la ficha principal" ,"Información");
				return;
			}
			if (cmbReporte.Text.Trim().Length == 0)
			{
				MessageBox.Show("Escoja un reporte" ,"Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			ParameterFields paramFields = new ParameterFields ();

			#region Parametro Fecha Desde
			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaDesde";
			discreteVal1.Value = dtDesde.ToString("yyyy-MM-dd HH:mm:ss");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
		
			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@FechaHasta";
			discreteVal2.Value = dtHasta.ToString("yyyy-MM-dd HH:mm:ss");
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);
			#endregion Parametro Fecha Desde

			ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			paramField3.ParameterFieldName = "@idCliente";
			int IdCliente = 0;
			if (ultraGrid1.ActiveRow != null && !chkTodoCliente.Checked)
				IdCliente = int.Parse(ultraGrid1.ActiveRow.Cells["idCliente"].Value.ToString());
			discreteVal3.Value = IdCliente;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

//			ParameterField paramField4 = new ParameterField ();
//			ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
//			paramField4.ParameterFieldName = "@Bodega";
//			discreteVal4.Value = spnBodega.Value;
//			paramField4.CurrentValues.Add (discreteVal4);
//			paramFields.Add (paramField4);

			string stReporte = cmbReporte.Text;
			Reporte miRep = new Reporte(stReporte);
			miRep.MdiParent = this.MdiParent;
			miRep.crVista.ParameterFieldInfo = paramFields;
			Cursor = Cursors.Default;
			miRep.Show();
		}

	}
}
