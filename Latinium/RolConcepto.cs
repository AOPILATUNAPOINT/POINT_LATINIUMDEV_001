using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolConcepto.
	/// </summary>
	public class RolConcepto : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRolConcepto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFormula;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel lblCondicion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCondicion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFormula;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnPosicion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCampo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraButton btGrupo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSQL;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImprimir;
		private Infragistics.Win.Misc.UltraButton btGenera;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkModificar;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private C1.Data.C1DataView cdvRolConcepto;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCodCuenta;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optDebeHaber;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo1;
		private C1.Data.C1DataView cdvCuentaDesc;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo3;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdRolConcepto;
		private Infragistics.Win.Misc.UltraButton btPosicion;
		private Infragistics.Win.Misc.UltraButton btGrilla;
		private Infragistics.Win.Misc.UltraButton btReordenar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnOrden;
		private DevExpress.XtraGrid.GridControl grdConceptos;
		private DevExpress.XtraGrid.Views.Grid.GridView grvConceptos;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolConcepto()
		{
			InitializeComponent();
		}

		int IdConcepto = 0;
		public RolConcepto(int idConcepto)
		{
			InitializeComponent();
			IdConcepto = idConcepto;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			this.cdsRolConcepto = new C1.Data.C1DataSet();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.chkFormula = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtFormula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.spnPosicion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCampo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCondicion = new Infragistics.Win.Misc.UltraLabel();
			this.txtCondicion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.btGrupo = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtSQL = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkImprimir = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btGenera = new Infragistics.Win.Misc.UltraButton();
			this.chkModificar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.spnOrden = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdvRolConcepto = new C1.Data.C1DataView();
			this.optDebeHaber = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbCodCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.ultraCombo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvCuentaDesc = new C1.Data.C1DataView();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCombo3 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.txtIdRolConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btPosicion = new Infragistics.Win.Misc.UltraButton();
			this.btGrilla = new Infragistics.Win.Misc.UltraButton();
			this.btReordenar = new Infragistics.Win.Misc.UltraButton();
			this.grdConceptos = new DevExpress.XtraGrid.GridControl();
			this.grvConceptos = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFormula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPosicion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCampo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCondicion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSQL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnOrden)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optDebeHaber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaDesc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdConceptos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvConceptos)).BeginInit();
			this.SuspendLayout();
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
			// txtConcepto
			// 
			this.txtConcepto.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtConcepto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.Detalle"));
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Location = new System.Drawing.Point(104, 17);
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(400, 22);
			this.txtConcepto.TabIndex = 1;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(51, 15);
			this.ultraLabel1.TabIndex = 2;
			this.ultraLabel1.Text = "Concepto";
			// 
			// chkFormula
			// 
			this.chkFormula.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkFormula.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsRolConcepto, "RolConcepto.Calculo"));
			this.chkFormula.Location = new System.Drawing.Point(32, 43);
			this.chkFormula.Name = "chkFormula";
			this.chkFormula.Size = new System.Drawing.Size(64, 22);
			this.chkFormula.TabIndex = 3;
			this.chkFormula.Text = "Formula";
			this.chkFormula.CheckedChanged += new System.EventHandler(this.chkFormula_CheckedChanged);
			// 
			// txtFormula
			// 
			this.txtFormula.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFormula.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.Formula"));
			this.txtFormula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFormula.Location = new System.Drawing.Point(104, 43);
			this.txtFormula.Multiline = true;
			this.txtFormula.Name = "txtFormula";
			this.txtFormula.ReadOnly = true;
			this.txtFormula.Size = new System.Drawing.Size(400, 60);
			this.txtFormula.TabIndex = 4;
			this.txtFormula.BeforeExitEditMode += new Infragistics.Win.BeforeExitEditModeEventHandler(this.txtFormula_BeforeExitEditMode);
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.idRolTipo"));
			this.cmbTipo.DataMember = "RolTipo";
			this.cmbTipo.DataSource = this.cdsRolConcepto;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(104, 138);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(104, 21);
			this.cmbTipo.TabIndex = 5;
			this.cmbTipo.ValueMember = "idRolTipo";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(32, 138);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(26, 15);
			this.ultraLabel2.TabIndex = 6;
			this.ultraLabel2.Text = "Tipo";
			// 
			// spnPosicion
			// 
			this.spnPosicion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.Posicion"));
			this.spnPosicion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnPosicion.Location = new System.Drawing.Point(104, 164);
			this.spnPosicion.Name = "spnPosicion";
			this.spnPosicion.PromptChar = ' ';
			this.spnPosicion.Size = new System.Drawing.Size(72, 22);
			this.spnPosicion.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnPosicion.TabIndex = 7;
			this.spnPosicion.Value = 1;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(32, 164);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel3.TabIndex = 8;
			this.ultraLabel3.Text = "Posicion";
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(232, 138);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel4.TabIndex = 10;
			this.ultraLabel4.Text = "Codigo";
			// 
			// txtCampo
			// 
			this.txtCampo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.Concepto"));
			this.txtCampo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCampo.Location = new System.Drawing.Point(296, 138);
			this.txtCampo.Name = "txtCampo";
			this.txtCampo.Size = new System.Drawing.Size(104, 22);
			this.txtCampo.TabIndex = 9;
			// 
			// lblCondicion
			// 
			this.lblCondicion.AutoSize = true;
			this.lblCondicion.Location = new System.Drawing.Point(32, 112);
			this.lblCondicion.Name = "lblCondicion";
			this.lblCondicion.Size = new System.Drawing.Size(53, 15);
			this.lblCondicion.TabIndex = 12;
			this.lblCondicion.Text = "Condicion";
			// 
			// txtCondicion
			// 
			this.txtCondicion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.Condicion"));
			this.txtCondicion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCondicion.Location = new System.Drawing.Point(104, 112);
			this.txtCondicion.Name = "txtCondicion";
			this.txtCondicion.ReadOnly = true;
			this.txtCondicion.Size = new System.Drawing.Size(344, 22);
			this.txtCondicion.TabIndex = 11;
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.idRolGrupo"));
			this.cmbGrupo.DataMember = "RolGrupo";
			this.cmbGrupo.DataSource = this.cdsRolConcepto;
			this.cmbGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 85;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(296, 164);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(104, 21);
			this.cmbGrupo.TabIndex = 13;
			this.cmbGrupo.ValueMember = "idRolGrupo";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(232, 164);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel5.TabIndex = 14;
			this.ultraLabel5.Text = "Grupo";
			// 
			// btGrupo
			// 
			this.btGrupo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrupo.Location = new System.Drawing.Point(416, 164);
			this.btGrupo.Name = "btGrupo";
			this.btGrupo.Size = new System.Drawing.Size(24, 24);
			this.btGrupo.TabIndex = 15;
			this.btGrupo.Text = "...";
			this.btGrupo.Click += new System.EventHandler(this.btGrupo_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtSQL
			// 
			this.txtSQL.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtSQL.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.FormulaSQL"));
			this.txtSQL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSQL.Location = new System.Drawing.Point(32, 319);
			this.txtSQL.Multiline = true;
			this.txtSQL.Name = "txtSQL";
			this.txtSQL.Size = new System.Drawing.Size(352, 137);
			this.txtSQL.TabIndex = 16;
			// 
			// chkImprimir
			// 
			this.chkImprimir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkImprimir.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsRolConcepto, "RolConcepto.Mostrar"));
			this.chkImprimir.Location = new System.Drawing.Point(344, 190);
			this.chkImprimir.Name = "chkImprimir";
			this.chkImprimir.Size = new System.Drawing.Size(64, 21);
			this.chkImprimir.TabIndex = 17;
			this.chkImprimir.Text = "Visible";
			// 
			// btGenera
			// 
			this.btGenera.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGenera.Location = new System.Drawing.Point(416, 138);
			this.btGenera.Name = "btGenera";
			this.btGenera.Size = new System.Drawing.Size(24, 25);
			this.btGenera.TabIndex = 18;
			this.btGenera.Text = "...";
			this.btGenera.Click += new System.EventHandler(this.btGenera_Click);
			this.btGenera.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btGenera_MouseUp);
			// 
			// chkModificar
			// 
			this.chkModificar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkModificar.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsRolConcepto, "RolConcepto.Modificar"));
			this.chkModificar.Location = new System.Drawing.Point(232, 190);
			this.chkModificar.Name = "chkModificar";
			this.chkModificar.Size = new System.Drawing.Size(72, 21);
			this.chkModificar.TabIndex = 19;
			this.chkModificar.Text = "Modificar";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(32, 190);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel6.TabIndex = 21;
			this.ultraLabel6.Text = "Orden";
			// 
			// spnOrden
			// 
			this.spnOrden.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.Orden"));
			this.spnOrden.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnOrden.Location = new System.Drawing.Point(104, 190);
			this.spnOrden.MaskInput = "nnn";
			this.spnOrden.Name = "spnOrden";
			this.spnOrden.PromptChar = ' ';
			this.spnOrden.Size = new System.Drawing.Size(72, 22);
			this.spnOrden.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnOrden.TabIndex = 20;
			this.spnOrden.Value = 1;
			// 
			// cdvRolConcepto
			// 
			this.cdvRolConcepto.BindingContextCtrl = this;
			this.cdvRolConcepto.DataSet = this.cdsRolConcepto;
			this.cdvRolConcepto.Sort = "idRolConcepto";
			this.cdvRolConcepto.TableName = "";
			this.cdvRolConcepto.TableViewName = "RolConcepto";
			// 
			// optDebeHaber
			// 
			this.optDebeHaber.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.DebeHaber"));
			this.optDebeHaber.ItemAppearance = appearance1;
			this.optDebeHaber.ItemOrigin = new System.Drawing.Point(6, 4);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Debe";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Haber";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Ambos";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "Ninguno";
			this.optDebeHaber.Items.Add(valueListItem1);
			this.optDebeHaber.Items.Add(valueListItem2);
			this.optDebeHaber.Items.Add(valueListItem3);
			this.optDebeHaber.Items.Add(valueListItem4);
			this.optDebeHaber.Location = new System.Drawing.Point(104, 215);
			this.optDebeHaber.Name = "optDebeHaber";
			this.optDebeHaber.Size = new System.Drawing.Size(280, 26);
			this.optDebeHaber.TabIndex = 22;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Location = new System.Drawing.Point(32, 250);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(39, 15);
			this.ultraLabel7.TabIndex = 24;
			this.ultraLabel7.Text = "Cuenta";
			// 
			// cmbCodCuenta
			// 
			this.cmbCodCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCodCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodCuenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.idCuenta"));
			this.cmbCodCuenta.DataSource = this.cdvCuenta;
			ultraGridBand3.AddButtonCaption = "RolTipo";
			ultraGridColumn5.Header.VisiblePosition = 8;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 8;
			ultraGridColumn6.Header.VisiblePosition = 12;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 8;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 8;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 8;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 8;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 8;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 17;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 17;
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 17;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 8;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 17;
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 17;
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 17;
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 17;
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Width = 146;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 287;
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn24});
			this.cmbCodCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCodCuenta.DisplayMember = "CodRapido";
			this.cmbCodCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCodCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCodCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodCuenta.Location = new System.Drawing.Point(112, 250);
			this.cmbCodCuenta.Name = "cmbCodCuenta";
			this.cmbCodCuenta.Size = new System.Drawing.Size(128, 21);
			this.cmbCodCuenta.TabIndex = 23;
			this.cmbCodCuenta.ValueMember = "idCuenta";
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo = 0";
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
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// ultraCombo1
			// 
			this.ultraCombo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.idCuenta"));
			this.ultraCombo1.DataSource = this.cdvCuentaDesc;
			ultraGridBand4.AddButtonCaption = "RolTipo";
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 8;
			ultraGridColumn26.Header.VisiblePosition = 12;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 8;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 139;
			ultraGridColumn28.Header.VisiblePosition = 9;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 8;
			ultraGridColumn29.Header.VisiblePosition = 5;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 8;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 8;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 17;
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 17;
			ultraGridColumn33.Header.VisiblePosition = 7;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 17;
			ultraGridColumn34.Header.VisiblePosition = 10;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 8;
			ultraGridColumn35.Header.VisiblePosition = 11;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 17;
			ultraGridColumn36.Header.VisiblePosition = 13;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 17;
			ultraGridColumn37.Header.VisiblePosition = 14;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 17;
			ultraGridColumn38.Header.VisiblePosition = 15;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 17;
			ultraGridColumn39.Header.VisiblePosition = 16;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 17;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 18;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 2;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 146;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Width = 287;
			ultraGridColumn44.Header.VisiblePosition = 19;
			ultraGridColumn44.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn44});
			this.ultraCombo1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraCombo1.DisplayMember = "Descripcion";
			this.ultraCombo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo1.Location = new System.Drawing.Point(248, 250);
			this.ultraCombo1.Name = "ultraCombo1";
			this.ultraCombo1.Size = new System.Drawing.Size(200, 21);
			this.ultraCombo1.TabIndex = 25;
			this.ultraCombo1.ValueMember = "idCuenta";
			this.ultraCombo1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ultraCombo1_MouseDown);
			// 
			// cdvCuentaDesc
			// 
			this.cdvCuentaDesc.BindingContextCtrl = this;
			this.cdvCuentaDesc.DataSet = this.cdsCuenta;
			this.cdvCuentaDesc.RowFilter = "Grupo = 0";
			this.cdvCuentaDesc.Sort = "Descripcion";
			this.cdvCuentaDesc.TableName = "";
			this.cdvCuentaDesc.TableViewName = "Cuenta";
			// 
			// ultraCombo2
			// 
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.idCuenta2"));
			this.ultraCombo2.DataSource = this.cdvCuentaDesc;
			ultraGridBand5.AddButtonCaption = "RolTipo";
			ultraGridColumn45.Header.VisiblePosition = 8;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 8;
			ultraGridColumn46.Header.VisiblePosition = 12;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 8;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 136;
			ultraGridColumn48.Header.VisiblePosition = 9;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 8;
			ultraGridColumn49.Header.VisiblePosition = 5;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 6;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 3;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 17;
			ultraGridColumn52.Header.VisiblePosition = 4;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 17;
			ultraGridColumn53.Header.VisiblePosition = 7;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 17;
			ultraGridColumn54.Header.VisiblePosition = 10;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 8;
			ultraGridColumn55.Header.VisiblePosition = 11;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 17;
			ultraGridColumn56.Header.VisiblePosition = 13;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 17;
			ultraGridColumn57.Header.VisiblePosition = 14;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 17;
			ultraGridColumn58.Header.VisiblePosition = 15;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 17;
			ultraGridColumn59.Header.VisiblePosition = 16;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 17;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 18;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 146;
			ultraGridColumn63.Header.VisiblePosition = 0;
			ultraGridColumn63.Width = 287;
			ultraGridColumn64.Header.VisiblePosition = 19;
			ultraGridColumn64.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn64});
			this.ultraCombo2.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ultraCombo2.DisplayMember = "Descripcion";
			this.ultraCombo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(248, 284);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(200, 21);
			this.ultraCombo2.TabIndex = 28;
			this.ultraCombo2.ValueMember = "idCuenta";
			this.ultraCombo2.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ultraCombo2_MouseUp);
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.Location = new System.Drawing.Point(32, 284);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(63, 15);
			this.ultraLabel8.TabIndex = 27;
			this.ultraLabel8.Text = "Haber Prov.";
			// 
			// ultraCombo3
			// 
			this.ultraCombo3.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.idCuenta2"));
			this.ultraCombo3.DataSource = this.cdvCuenta;
			ultraGridBand6.AddButtonCaption = "RolTipo";
			ultraGridColumn65.Header.VisiblePosition = 8;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 8;
			ultraGridColumn66.Header.VisiblePosition = 12;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 8;
			ultraGridColumn67.Header.VisiblePosition = 2;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 8;
			ultraGridColumn68.Header.VisiblePosition = 9;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 8;
			ultraGridColumn69.Header.VisiblePosition = 5;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 8;
			ultraGridColumn70.Header.VisiblePosition = 6;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 8;
			ultraGridColumn71.Header.VisiblePosition = 3;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 17;
			ultraGridColumn72.Header.VisiblePosition = 4;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 17;
			ultraGridColumn73.Header.VisiblePosition = 7;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 17;
			ultraGridColumn74.Header.VisiblePosition = 10;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 8;
			ultraGridColumn75.Header.VisiblePosition = 11;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 17;
			ultraGridColumn76.Header.VisiblePosition = 13;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 17;
			ultraGridColumn77.Header.VisiblePosition = 14;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 17;
			ultraGridColumn78.Header.VisiblePosition = 15;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 17;
			ultraGridColumn79.Header.VisiblePosition = 16;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 17;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 18;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 0;
			ultraGridColumn82.Width = 146;
			ultraGridColumn83.Header.VisiblePosition = 1;
			ultraGridColumn83.Width = 287;
			ultraGridColumn84.Header.VisiblePosition = 19;
			ultraGridColumn84.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn84});
			this.ultraCombo3.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraCombo3.DisplayMember = "CodRapido";
			this.ultraCombo3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo3.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo3.Location = new System.Drawing.Point(112, 284);
			this.ultraCombo3.Name = "ultraCombo3";
			this.ultraCombo3.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo3.TabIndex = 26;
			this.ultraCombo3.ValueMember = "idCuenta";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "RolConcepto";
			this.barraDato1.DataNombreId = "idRolConcepto";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsRolConcepto;
			this.barraDato1.DataTabla = "RolConcepto";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(32, 471);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 25);
			this.barraDato1.TabIndex = 29;
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
			this.barraDato1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
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
			// txtIdRolConcepto
			// 
			this.txtIdRolConcepto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolConcepto, "RolConcepto.idRolConcepto"));
			this.txtIdRolConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRolConcepto.Location = new System.Drawing.Point(464, 472);
			this.txtIdRolConcepto.Name = "txtIdRolConcepto";
			this.txtIdRolConcepto.Size = new System.Drawing.Size(48, 22);
			this.txtIdRolConcepto.TabIndex = 30;
			// 
			// btPosicion
			// 
			this.btPosicion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPosicion.Location = new System.Drawing.Point(184, 164);
			this.btPosicion.Name = "btPosicion";
			this.btPosicion.Size = new System.Drawing.Size(24, 24);
			this.btPosicion.TabIndex = 31;
			this.btPosicion.Text = "...";
			this.btPosicion.Click += new System.EventHandler(this.btPosicion_Click);
			// 
			// btGrilla
			// 
			this.btGrilla.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrilla.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrilla.Location = new System.Drawing.Point(344, 471);
			this.btGrilla.Name = "btGrilla";
			this.btGrilla.Size = new System.Drawing.Size(64, 24);
			this.btGrilla.TabIndex = 33;
			this.btGrilla.Text = "Refresca";
			this.btGrilla.Click += new System.EventHandler(this.btGrilla_Click);
			this.btGrilla.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btGrilla_MouseUp);
			// 
			// btReordenar
			// 
			this.btReordenar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btReordenar.Location = new System.Drawing.Point(184, 190);
			this.btReordenar.Name = "btReordenar";
			this.btReordenar.Size = new System.Drawing.Size(24, 24);
			this.btReordenar.TabIndex = 34;
			this.btReordenar.Text = "...";
			this.btReordenar.Click += new System.EventHandler(this.btReordenar_Click);
			this.btReordenar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btReordenar_MouseUp);
			// 
			// grdConceptos
			// 
			this.grdConceptos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdConceptos.EmbeddedNavigator
			// 
			this.grdConceptos.EmbeddedNavigator.Name = "";
			this.grdConceptos.Location = new System.Drawing.Point(384, 320);
			this.grdConceptos.MainView = this.grvConceptos;
			this.grdConceptos.Name = "grdConceptos";
			this.grdConceptos.Size = new System.Drawing.Size(136, 136);
			this.grdConceptos.TabIndex = 35;
			this.grdConceptos.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvConceptos});
			this.grdConceptos.DoubleClick += new System.EventHandler(this.grdConceptos_DoubleClick);
			// 
			// grvConceptos
			// 
			this.grvConceptos.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																												this.gridColumn1});
			this.grvConceptos.GridControl = this.grdConceptos;
			this.grvConceptos.Name = "grvConceptos";
			this.grvConceptos.OptionsView.ShowGroupPanel = false;
			// 
			// gridColumn1
			// 
			this.gridColumn1.Caption = "Concepto";
			this.gridColumn1.FieldName = "Codigos";
			this.gridColumn1.Name = "gridColumn1";
			this.gridColumn1.OptionsColumn.AllowEdit = false;
			this.gridColumn1.OptionsColumn.AllowIncrementalSearch = false;
			this.gridColumn1.OptionsColumn.ReadOnly = true;
			this.gridColumn1.OptionsColumn.ShowCaption = false;
			this.gridColumn1.Visible = true;
			this.gridColumn1.VisibleIndex = 0;
			// 
			// RolConcepto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(528, 510);
			this.Controls.Add(this.grdConceptos);
			this.Controls.Add(this.btReordenar);
			this.Controls.Add(this.btGrilla);
			this.Controls.Add(this.btPosicion);
			this.Controls.Add(this.txtIdRolConcepto);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraCombo2);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraCombo3);
			this.Controls.Add(this.ultraCombo1);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.cmbCodCuenta);
			this.Controls.Add(this.optDebeHaber);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.spnOrden);
			this.Controls.Add(this.chkModificar);
			this.Controls.Add(this.btGenera);
			this.Controls.Add(this.chkImprimir);
			this.Controls.Add(this.txtSQL);
			this.Controls.Add(this.btGrupo);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.lblCondicion);
			this.Controls.Add(this.txtCondicion);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtCampo);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.spnPosicion);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.txtFormula);
			this.Controls.Add(this.chkFormula);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtConcepto);
			this.Name = "RolConcepto";
			this.Text = "Rubros de Roles de Pago";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolConcepto_Closing);
			this.Load += new System.EventHandler(this.RolConcepto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFormula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPosicion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCampo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCondicion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSQL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnOrden)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optDebeHaber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaDesc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdConceptos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvConceptos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private void UnloadMe()
		{
			this.Close();
		}
		private void RolConcepto_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("84",cdsSeguridad))
			{
				MessageBox.Show("No puede ingresar a Rubros", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			txtIdRolConcepto.Width = 0;

			grdConceptos.DataSource = Funcion.dvProcedimiento(cdsCuenta, "Select Distinct(Concepto) As Codigos From RolConcepto Order by Concepto");

			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Concepto'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			Cursor = Cursors.Default;
			miAcceso =new Acceso("0506", this.cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			if (IdConcepto > 0)
			{
				this.barraDato1.IdRegistro=IdConcepto;
				barraDato1.ProximoId(5);
			}
			else
			{
				this.barraDato1.PosUltima();
			}
			this.barraDato1.HabilitaControles(false);
		}

		private void chkFormula_CheckedChanged(object sender, System.EventArgs e)
		{
			this.txtFormula.ReadOnly = !this.chkFormula.Checked;
			this.txtCondicion.ReadOnly = !this.chkFormula.Checked;
		}

		private void btGrupo_Click(object sender, System.EventArgs e)
		{
			using (RolGrupo miGrupo = new RolGrupo())
			{
				miGrupo.ShowDialog();
			}
			this.cdsRolConcepto.Clear();
			this.cdsRolConcepto.Fill();
		}
		private Acceso miAcceso;

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Rubros de Roles", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Anular el Registro?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Buscar)
			{
				MessageBox.Show("No tiene acceso a Buscar Rubros de Roles", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			using (BuscaRolConcepto miBusqueda = new BuscaRolConcepto())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.idBusca.Text.Length > 0)
						RolBuscado(int.Parse(miBusqueda.idBusca.Text));
				}
			}
		}

		private void RolBuscado(int id)
		{
			barraDato1.IdRegistro = id;
			this.barraDato1.ProximoId(5);
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Crear nuevos Rubros de Roles", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.CrearRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Rubros de Roles", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.EditarRegistro();
		}

		private void BorraErrores()
		{
			errorProvider.SetError(cmbGrupo, "");
			errorProvider.SetError(cmbTipo, "");
			errorProvider.SetError(txtConcepto, "");
			errorProvider.SetError(txtCampo, "");
			errorProvider.SetError(spnPosicion, "");
			errorProvider.SetError(chkFormula, "");
			errorProvider.SetError(txtCampo, "");
		}
		private bool Verifica()
		{
			bool bOk = true;
			this.BorraErrores();
			if (chkFormula.Checked && txtFormula.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.chkFormula, "No existe una formula definida");
				bOk = false;
			}
			if (this.cmbGrupo.Value == null || (int) this.cmbGrupo.Value == 0)
			{
				this.errorProvider.SetError(this.cmbGrupo, "Ingrese Grupo");
				bOk = false;
			}
			if (this.cmbTipo.Value == null || (int) this.cmbTipo.Value == 0)
			{
				this.errorProvider.SetError(this.cmbTipo, "Ingrese Tipo");
				bOk = false;
			}
			if (this.txtConcepto.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtConcepto, "Ingrese Concepto");
				bOk = false;
			}
			if (this.txtCampo.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtCampo, "Ingrese Codigo");
				bOk = false;
			}
			if ((int) this.spnPosicion.Value < 1)
			{
				this.errorProvider.SetError(this.spnPosicion, "Ingrese Posicion");
				bOk = false;
			}
			string stCampo = txtCampo.Text.Trim();
			if ((int) stCampo.IndexOf(" ") > 0)
			{
				errorProvider.SetError(txtCampo, "Codigo no puede tener espacios vacios");
				bOk = false;
			}
			return bOk;
		}

		private bool VerificaSql()
		{
			int iTipo = (int)cmbTipo.Value;
			int iPos = (int)spnPosicion.Value;
			int iGrupo = (int)cmbGrupo.Value;

			string stConcepto = "Select Count(*) From RolConcepto r Where r.idRolGrupo = " + iGrupo;
			stConcepto += " And r.idRolTipo = " + iTipo + " And r.Posicion = " + iPos;
			if (!barraDato1.bNuevo) 
        stConcepto += " And r.idRolConcepto <> " + txtIdRolConcepto.Text;

			int iCuenta = Funcion.iEscalarSQL(cdsCuenta, stConcepto, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("Tipo y Posicion ya existen para ese grupo", "Información");
				return false;
			}
			string stExistePersonal = "";
			if (txtSQL.Text.ToLower().IndexOf("personal.") >= 0 || txtCondicion.Text.ToLower().IndexOf("personal.") >= 0)
				stExistePersonal = "Inner Join Personal On RolPago.idPersonal = Personal.idPersonal";

			if (chkFormula.Checked)
			{
				string stSelect = string.Format("Select {0} from RolPago {1}"
					, txtSQL.Text.Trim(), stExistePersonal);
				if (txtCondicion.Text.Trim().Length > 0) stSelect += " Where " + txtCondicion.Text.Trim();
				string stMensaje = Funcion.VerificaSQL(cdsCuenta, stSelect);
				if (stMensaje.Length > 0)
				{
					MessageBox.Show("Formula generada incorrecta. Verifique nombres de campos.\n" + stMensaje, stSelect);
					return false;
				}
			}

			return true;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			if (chkFormula.Checked) btGenera_Click(this, e);
			if (!VerificaSql()) return;
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

		private void btGenera_Click(object sender, System.EventArgs e)
		{
			if (this.txtFormula.Text.Length == 0) return;
			string stConcepto = this.txtFormula.Value.ToString();
			int iGrupo = (int) this.cmbGrupo.Value;
			stConcepto = stConcepto.Trim();
			char [] cSep = new char[7];
			char [] cOper = new char[200];
			cSep[0] = '+';
			cSep[1] = '-';
			cSep[2] = '*';
			cSep[3] = '/';
			cSep[4] = '(';
			cSep[5] = ')';
			cSep[6] = ' ';
			int k = 0;
			int i;
			for (i = 0; i< stConcepto.Length; i++)
			{
				for (int j =0; j< 7; j++)
				{
					if (stConcepto[i] == cSep[j])
						cOper[k++] = cSep[j];
				}
			}

			string stFormula = "";
			i=0;
			foreach (string stConc in	stConcepto.Split(cSep))
			{
				if (i!=0) stFormula += cOper[i-1];
				stFormula += Funcion.RolCampo(cdsRolConcepto, stConc, iGrupo);
				i++;
			}
			this.txtSQL.Value = stFormula;
		}

		private void cdsRolConcepto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolConcepto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
			RolConceptoCopia miCopia = new RolConceptoCopia();
			miCopia.MdiParent = this.MdiParent;
			miCopia.Show();
		}

		private void RolConcepto_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Concepto'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}

		private void ultraCombo1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			ultraCombo1.Value = 0;
			cmbCodCuenta.Value = 0;
			cmbCodCuenta.Select();
			cmbCodCuenta.Refresh();

		}

		private void ultraCombo2_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			ultraCombo2.Value = 0;
			ultraCombo3.Value = 0;
			ultraCombo3.Select();
			ultraCombo3.Refresh();
		}

		private void btPosicion_Click(object sender, System.EventArgs e)
		{
			int iTipo = (int)cmbTipo.Value;
			int iGrupo = (int)cmbGrupo.Value;

			string stConcepto = "Select Max(Posicion) From RolConcepto r Where r.idRolGrupo = " + iGrupo +
				" And r.idRolTipo = " + iTipo;
			int iPos = 1;
			try
			{
				iPos = Funcion.iEscalarSQL(cdsCuenta, stConcepto, true) + 1;
			}
			catch{};
			spnPosicion.Value = iPos;
			spnPosicion.SelectAll();
			spnPosicion.Refresh();
		}

		private void btReordenar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea crear un espacio despues de este orden?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			int iGrupo = (int)cmbGrupo.Value;
			int iOrden = (int) spnOrden.Value;

			string stConcepto = "Update RolConcepto Set Orden = Orden + 1 Where idRolGrupo = " + iGrupo +
				" And Orden > " + iOrden;
			try
			{
				Funcion.EjecutaSQL(cdsCuenta, stConcepto, true);
			}
			catch{};

			MessageBox.Show("Espacio creado. Puede insertar un rubro en posicion " + (iOrden + 1).ToString(), 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btReordenar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Reordenar los rubros?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			int iGrupo = (int)cmbGrupo.Value;
			int iOrden = (int) spnOrden.Value;
			string stConcepto = "Exec RolReordena " + iGrupo + ", " + iOrden;
			try
			{
				Funcion.EjecutaSQL(cdsCuenta, stConcepto, true);
			}
			catch{};
			barraDato1.ProximoId(5);
			MessageBox.Show("Rubros Reordenados", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btGenera_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Regenerar Todas las Fórmulas de este grupo?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			int iGrupo = (int) this.cmbGrupo.Value;
			char [] cSep = new char[7];
			char [] cOper = new char[200];
			cSep[0] = '+';
			cSep[1] = '-';
			cSep[2] = '*';
			cSep[3] = '/';
			cSep[4] = '(';
			cSep[5] = ')';
			cSep[6] = ' ';
			int i;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsCuenta, 
				"Select Formula, idRolConcepto From RolConcepto Where idRolGrupo = " 
				+ iGrupo + " And Calculo = 1 And Len(IsNull(FormulaSQL, '')) > 0", true);
			string[] stFormulas = new string [1000];
			int[] idRolConcepto = new int [1000];
			int m = 0;
			while(dr.Read())
			{
				stFormulas[m] = dr.GetString(0);
				idRolConcepto[m++] = dr.GetInt32(1);
			}
			dr.Close();

			for (int l=0; l<m;l++)
			{
				string stConcepto = stFormulas[l];
				stConcepto = stConcepto.Trim();
				int k = 0;
				for (i = 0; i< stConcepto.Length; i++)
				{
					for (int j =0; j< 7; j++)
					{
						if (stConcepto[i] == cSep[j])
							cOper[k++] = cSep[j];
					}
				}

				string stFormula = "";
				i=0;
				foreach (string stConc in	stConcepto.Split(cSep))
				{
					if (i!=0) stFormula += cOper[i-1];
					stFormula += Funcion.RolCampo(cdsRolConcepto, stConc, iGrupo);
					i++;
				}
				string stSelect = string.Format("Update RolConcepto Set FormulaSQL = '{0}' Where idRolConcepto = {1}",
					stFormula, idRolConcepto[l]) ;
				Funcion.EjecutaSQL(cdsCuenta, stSelect);
			}
			MessageBox.Show("Todos las Formulas estan regeneradas", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btGrilla_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show("¿Desea Verificar Todas las Fórmulas de este grupo?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stMesajes ="";
			SqlDataReader dr = Funcion.rEscalarSQL(cdsCuenta, 
				"Select FormulaSQL, Concepto, IsNull(Condicion, '') From RolConcepto Where Calculo = 1 And Len(IsNull(FormulaSQL, '')) > 0", true);
			int i = 0;
			string[] stFormulas = new string[1000];
			string[] stConceptos = new string[1000];
			string[] stCondicion = new string[1000];
			while (dr.Read())
			{
				stFormulas[i] = dr.GetString(0).Trim();
				stConceptos[i] = dr.GetString(1).Trim();
				stCondicion[i++] = dr.GetString(2).Trim();
			}
			dr.Close();

			int k= 0;
			for (int j = 0;j<i;j++)
			{
				string stExistePersonal = "";
				if (stFormulas[j].ToLower().IndexOf("personal.") >= 0 || stCondicion[j].ToLower().IndexOf("personal.") >= 0)
					stExistePersonal = "Inner Join Personal On RolPago.idPersonal = Personal.idPersonal";
				string stSelect = string.Format("Select {0} from RolPago {1}"
					, stFormulas[j].Trim(), stExistePersonal);
				if (stCondicion[j].Length > 0) stSelect += " Where " +  stCondicion[j];
				string stMensaje = Funcion.VerificaSQL(cdsCuenta, stSelect);
				if (stMensaje.Length > 0)
				{
					k++;
					stMesajes += "Codigo: " + stConceptos[j] + " Error: " + stMensaje + "\n";
				}
			}
			if (k > 0)
			{
				MessageBox.Show("Rubros con errores:\n" + stMesajes, 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
			else
			{
				MessageBox.Show("Todos los Rubros estan perfectos", 
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}

		}

		int iPosCursor = 0;
		int iSelLength = 0;
		private void txtFormula_BeforeExitEditMode(object sender, Infragistics.Win.BeforeExitEditModeEventArgs e)
		{
			iPosCursor = txtFormula.SelectionStart;
			iSelLength = txtFormula.SelectionLength;
		}

		private void grdConceptos_DoubleClick(object sender, System.EventArgs e)
		{
			int iFila = grvConceptos.FocusedRowHandle;
			if (iFila < 0) return;

			string stConcepto = grvConceptos.GetRowCellValue(iFila, "Codigos").ToString();
			string stFormula = txtFormula.Text.Trim();
			try
			{
				if (iSelLength > 0)
				{
					stFormula = stFormula.Remove(iPosCursor, iSelLength);
					iSelLength = 0;
				}
				if (iPosCursor <= stFormula.Length)
				{
					txtFormula.Text = stFormula.Insert(iPosCursor, stConcepto);
					txtFormula.Focus();
					iPosCursor += stConcepto.Length;
					txtFormula.SelectionLength = 0;
					txtFormula.SelectionStart = iPosCursor;
				}
				else
					iPosCursor = stFormula.Length;
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error: Insertar Codigo");
			}
		}

		private void btGrilla_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Refrescar grilla con nuevos conceptos generados?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			grdConceptos.DataSource = Funcion.dvProcedimiento(cdsCuenta, 
				"Select Concepto As Codigos From RolConcepto Order by Concepto");

		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
