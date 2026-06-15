using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Presupuesto.
	/// </summary>
	public class Presupuesto : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsPresupuesto;
		private C1.Data.C1DataSet cdsCuenta;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodCuenta;
		private C1.Data.C1DataView cdvCodCuenta;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private C1.Data.C1DataSet cdsProyecto;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyLinea;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubPLinea;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAnual;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.Label label6;
		private C1.Data.C1DataSet cdsSeguridad;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaNombre;
		private C1.Data.C1DataView cdvCuentaNombre;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalH;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalD;
		private System.Windows.Forms.Button btTotal;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Presupuesto()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_PresupuestoDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCodCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyLinea");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto", -1, "cmbSubPLinea");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaPresup", 0, "cmbCuentaNombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPuntos");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.cdsPresupuesto = new C1.Data.C1DataSet();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbCodCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvCodCuenta = new C1.Data.C1DataView();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbProyLinea = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cmbSubPLinea = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkAnual = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.label6 = new System.Windows.Forms.Label();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cdvCuentaNombre = new C1.Data.C1DataView();
			this.cmbCuentaNombre = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtTotalH = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalD = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btTotal = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyLinea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubPLinea)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalD)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsPresupuesto
			// 
			this.cdsPresupuesto.BindingContextCtrl = this;
			this.cdsPresupuesto.DataLibrary = "LibContabilidad";
			this.cdsPresupuesto.DataLibraryUrl = "";
			this.cdsPresupuesto.DataSetDef = "dsPresupuesto";
			this.cdsPresupuesto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPresupuesto.Name = "cdsPresupuesto";
			this.cdsPresupuesto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsPresupuesto.SchemaDef = null;
			this.cdsPresupuesto.BeforeFill += new C1.Data.FillEventHandler(this.cdsPresupuesto_BeforeFill);
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Presupuesto.Presupuesto - PresupuestoDetalle";
			this.ultraGrid1.DataSource = this.cdsPresupuesto;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "_Presupuesto";
			ultraGridColumn1.Header.VisiblePosition = 4;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 76;
			ultraGridColumn2.Header.VisiblePosition = 6;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 72;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 78;
			ultraGridColumn4.Header.Caption = "Codigo P.";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn4.Width = 83;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 84;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 82;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 69;
			ultraGridColumn9.Header.Caption = "C. Costo";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn9.Width = 34;
			ultraGridColumn10.Header.Caption = "Proyecto";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 68;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 71;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 89;
			ultraGridColumn13.Header.Caption = "Descripcion";
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn13.Width = 160;
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
																										 ultraGridColumn13});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 95);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(632, 198);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cmbCodCuenta
			// 
			this.cmbCodCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodCuenta.DataSource = this.cdvCodCuenta;
			ultraGridColumn14.Header.VisiblePosition = 8;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 9;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 112;
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 7;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 10;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 11;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 13;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 14;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 15;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 16;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 17;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 18;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Width = 384;
			ultraGridColumn33.Header.VisiblePosition = 19;
			ultraGridColumn33.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn33});
			this.cmbCodCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCodCuenta.DisplayMember = "CodRapido";
			this.cmbCodCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodCuenta.Location = new System.Drawing.Point(240, 121);
			this.cmbCodCuenta.Name = "cmbCodCuenta";
			this.cmbCodCuenta.Size = new System.Drawing.Size(272, 94);
			this.cmbCodCuenta.TabIndex = 1;
			this.cmbCodCuenta.ValueMember = "idCuenta";
			this.cmbCodCuenta.Visible = false;
			// 
			// cdvCodCuenta
			// 
			this.cdvCodCuenta.BindingContextCtrl = this;
			this.cdvCodCuenta.DataSet = this.cdsCuenta;
			this.cdvCodCuenta.RowFilter = "CodPresup > \'\'";
			this.cdvCodCuenta.Sort = "Codigo";
			this.cdvCodCuenta.TableName = "";
			this.cdvCodCuenta.TableViewName = "Cuenta";
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "_Presupuesto";
			this.barraDatoSQL1.DataNombreId = "idPresupuesto";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsPresupuesto;
			this.barraDatoSQL1.DataTabla = "Presupuesto";
			this.barraDatoSQL1.DataTablaHija = "PresupuestoDetalle";
			this.barraDatoSQL1.Location = new System.Drawing.Point(16, 310);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(280, 26);
			this.barraDatoSQL1.TabIndex = 2;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPresupuesto, "_Presupuesto.Detalle"));
			this.txtDescripcion.Location = new System.Drawing.Point(288, 17);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(264, 43);
			this.txtDescripcion.TabIndex = 3;
			// 
			// cmbFecha
			// 
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPresupuesto, "_Presupuesto.Fecha"));
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(72, 43);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(112, 21);
			this.cmbFecha.TabIndex = 4;
			this.cmbFecha.Value = new System.DateTime(2005, 11, 2, 0, 0, 0, 0);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPresupuesto, "_Presupuesto.idProyecto"));
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Location = new System.Drawing.Point(72, 69);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(112, 22);
			this.cmbProyecto.TabIndex = 5;
			this.cmbProyecto.ValueMember = "idProyecto";
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
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 69);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 17);
			this.label1.TabIndex = 6;
			this.label1.Text = "C. Costo";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(224, 69);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 17);
			this.label2.TabIndex = 7;
			this.label2.Text = "Proyecto";
			// 
			// ultraCombo2
			// 
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPresupuesto, "_Presupuesto.idSubProyecto"));
			this.ultraCombo2.DataMember = "SubProyecto";
			this.ultraCombo2.DataSource = this.cdsProyecto;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn37.Header.VisiblePosition = 3;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 6;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 5;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 7;
			ultraGridColumn41.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.ultraCombo2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraCombo2.DisplayMember = "Nombre";
			this.ultraCombo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(288, 69);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(112, 21);
			this.ultraCombo2.TabIndex = 8;
			this.ultraCombo2.ValueMember = "idSubProyecto";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 17);
			this.label3.TabIndex = 9;
			this.label3.Text = "Fecha";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(224, 17);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(56, 17);
			this.label4.TabIndex = 10;
			this.label4.Text = "Detalle";
			// 
			// cmbProyLinea
			// 
			this.cmbProyLinea.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyLinea.DataSource = this.cdvProyecto;
			ultraGridColumn42.Header.VisiblePosition = 2;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 6;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 0;
			ultraGridColumn44.Width = 210;
			ultraGridColumn45.Header.VisiblePosition = 3;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 7;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 4;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 127;
			ultraGridColumn49.Header.VisiblePosition = 5;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 8;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 9;
			ultraGridColumn51.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51});
			this.cmbProyLinea.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbProyLinea.DisplayMember = "Nombre";
			this.cmbProyLinea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyLinea.Location = new System.Drawing.Point(56, 144);
			this.cmbProyLinea.Name = "cmbProyLinea";
			this.cmbProyLinea.Size = new System.Drawing.Size(272, 103);
			this.cmbProyLinea.TabIndex = 11;
			this.cmbProyLinea.ValueMember = "idProyecto";
			this.cmbProyLinea.Visible = false;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible = 1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cmbSubPLinea
			// 
			this.cmbSubPLinea.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubPLinea.DataSource = this.cdvSubProyecto;
			ultraGridColumn52.Header.VisiblePosition = 4;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 2;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 0;
			ultraGridColumn54.Width = 198;
			ultraGridColumn55.Header.VisiblePosition = 3;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 6;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 5;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 1;
			ultraGridColumn59.Header.VisiblePosition = 7;
			ultraGridColumn59.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59});
			this.cmbSubPLinea.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbSubPLinea.DisplayMember = "Nombre";
			this.cmbSubPLinea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubPLinea.Location = new System.Drawing.Point(216, 207);
			this.cmbSubPLinea.Name = "cmbSubPLinea";
			this.cmbSubPLinea.Size = new System.Drawing.Size(216, 86);
			this.cmbSubPLinea.TabIndex = 12;
			this.cmbSubPLinea.ValueMember = "idSubProyecto";
			this.cmbSubPLinea.Visible = false;
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.RowFilter = "Visible = 1";
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			// 
			// txtCodigo
			// 
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPresupuesto, "_Presupuesto.Codigo"));
			this.txtCodigo.Location = new System.Drawing.Point(72, 17);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 22);
			this.txtCodigo.TabIndex = 15;
			// 
			// chkAnual
			// 
			this.chkAnual.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.chkAnual.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPresupuesto, "_Presupuesto.Distribucion"));
			this.chkAnual.Location = new System.Drawing.Point(472, 64);
			this.chkAnual.Name = "chkAnual";
			this.chkAnual.Size = new System.Drawing.Size(152, 22);
			this.chkAnual.TabIndex = 16;
			this.chkAnual.Text = "Distribución Asignación";
			this.chkAnual.ThreeState = true;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(16, 17);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(48, 17);
			this.label6.TabIndex = 17;
			this.label6.Text = "Código";
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
			// cdvCuentaNombre
			// 
			this.cdvCuentaNombre.BindingContextCtrl = this;
			this.cdvCuentaNombre.DataSet = this.cdsCuenta;
			this.cdvCuentaNombre.RowFilter = "CodPresup > \'\'";
			this.cdvCuentaNombre.Sort = "Descripcion";
			this.cdvCuentaNombre.TableName = "";
			this.cdvCuentaNombre.TableViewName = "Cuenta";
			// 
			// cmbCuentaNombre
			// 
			this.cmbCuentaNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaNombre.DataSource = this.cdvCuentaNombre;
			ultraGridColumn60.Header.VisiblePosition = 3;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 4;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 125;
			ultraGridColumn63.Header.VisiblePosition = 5;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 6;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 8;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 9;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 10;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 11;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 12;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 7;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 13;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 14;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 15;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 16;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 17;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 18;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 1;
			ultraGridColumn77.Width = 111;
			ultraGridColumn78.Header.VisiblePosition = 0;
			ultraGridColumn78.Width = 310;
			ultraGridColumn79.Header.VisiblePosition = 19;
			ultraGridColumn79.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn79});
			this.cmbCuentaNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCuentaNombre.DisplayMember = "Descripcion";
			this.cmbCuentaNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaNombre.Location = new System.Drawing.Point(432, 224);
			this.cmbCuentaNombre.Name = "cmbCuentaNombre";
			this.cmbCuentaNombre.Size = new System.Drawing.Size(136, 69);
			this.cmbCuentaNombre.TabIndex = 18;
			this.cmbCuentaNombre.ValueMember = "idCuenta";
			this.cmbCuentaNombre.Visible = false;
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
			// txtTotalH
			// 
			this.txtTotalH.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalH.Appearance = appearance8;
			this.txtTotalH.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalH.FormatString = "#,##0.00";
			this.txtTotalH.Location = new System.Drawing.Point(488, 312);
			this.txtTotalH.Name = "txtTotalH";
			this.txtTotalH.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalH.PromptChar = ' ';
			this.txtTotalH.Size = new System.Drawing.Size(80, 22);
			this.txtTotalH.TabIndex = 130;
			// 
			// txtTotalD
			// 
			this.txtTotalD.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalD.Appearance = appearance9;
			this.txtTotalD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalD.FormatString = "#,##0.00";
			this.txtTotalD.Location = new System.Drawing.Point(400, 312);
			this.txtTotalD.Name = "txtTotalD";
			this.txtTotalD.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalD.PromptChar = ' ';
			this.txtTotalD.Size = new System.Drawing.Size(80, 22);
			this.txtTotalD.TabIndex = 133;
			// 
			// btTotal
			// 
			this.btTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btTotal.Location = new System.Drawing.Point(320, 312);
			this.btTotal.Name = "btTotal";
			this.btTotal.TabIndex = 134;
			this.btTotal.Text = "Total";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// Presupuesto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(664, 359);
			this.Controls.Add(this.btTotal);
			this.Controls.Add(this.txtTotalD);
			this.Controls.Add(this.txtTotalH);
			this.Controls.Add(this.cmbCuentaNombre);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.chkAnual);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.cmbProyLinea);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.ultraCombo2);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.cmbCodCuenta);
			this.Controls.Add(this.cmbSubPLinea);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "Presupuesto";
			this.Text = "Presupuesto";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Presupuesto_Closing);
			this.Load += new System.EventHandler(this.Presupuesto_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyLinea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubPLinea)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalD)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		string controlW="";
		string NuevoCodigo="";
		string controlW1="";
		
		private void cdsPresupuesto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPresupuesto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Presupuesto",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Desea Borrar Presupuesto?"
				, "Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			barraDatoSQL1.BorraRegistro();
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void CopiaPresupuesto(DateTime dtFecha, int iDivide)
		{
			string stInsert = "";
			string stNumero = Contabilidad.NumeroAsiento(cdsCuenta, 8);
			stInsert = "Insert into presupuesto "
				+ "(idSucursal, Detalle, Numero, "
				+ "Notas, idAsiento, Codigo, Distribucion, "
				+ "Borrar, idProyecto, idSubProyecto, Fecha) "
				+ "SELECT idSucursal, Detalle, Numero, "
				+ "Notas, idAsiento, '" + stNumero + "', 0, "
				+ "Borrar, idProyecto, idSubProyecto, '" + dtFecha.ToString("yyyyMMdd")
				+ "' FROM Presupuesto WHERE idPresupuesto = " 
				+ drPres["idPresupuesto"].ToString();
			Funcion.EjecutaSQL(cdsPresupuesto, stInsert);

			stInsert = "Select TOP 1 idPresupuesto FROM Presupuesto "
				+ "ORDER BY idPresupuesto DESC";
			int IdPres = Funcion.iEscalarSQL(cdsPresupuesto, stInsert);
			stInsert = "Insert into PresupuestoDetalle "
				+ "(idPresupuesto, Debe, idSucursal, Haber, "
				+ "idCuenta, Detalle, idProyecto, idSubProyecto) "
				+ "SELECT " + IdPres.ToString()
				+ ", Debe/" + iDivide.ToString()+  ", idSucursal, Haber/"
        + iDivide.ToString() + ", "
				+ "idCuenta, Detalle, idProyecto, idSubProyecto "
				+ "FROM PresupuestoDetalle Where idPresupuesto = " 
				+ drPres["idPresupuesto"].ToString();
			Funcion.EjecutaSQL(cdsPresupuesto, stInsert);
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Copiar o Crear Presupuesto",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if ((bool) chkAnual.Checked)
			{
				if (cmbFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha", "Información");
					return;
				}
				if (DialogResult.No == MessageBox.Show(
					"¿Desea Distribuir el presupuesto anual a todos los meses?",
					"Confirmación", MessageBoxButtons.YesNo, 
					MessageBoxIcon.Question)) return;
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				dtFecha = dtFecha.AddMonths(-dtFecha.Month); // Empieza en Enero
				for (int i=0; i<12; i++) 
				{
					dtFecha = dtFecha.AddMonths(1);
					CopiaPresupuesto(dtFecha, 12);
				}
				barraDatoSQL1.PosUltima();
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Desea Copiar presupuesto?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question)) return;
			CopiaPresupuesto(DateTime.Today, 1);
			barraDatoSQL1.PosUltima();
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Copiar o Crear Presupuesto",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			barraDatoSQL1.CrearRegistro();
			ModificaGrilla(true);
			drPres["Fecha"] = DateTime.Today;
			if (bNumera)
			{
				txtDescripcion.Select();
				txtCodigo.ReadOnly = true;
			}
			else
				txtCodigo.Select();
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			BorraErrores();
			barraDatoSQL1.DeshacerRegistro();
			ModificaGrilla(false);
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Presupuesto",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			barraDatoSQL1.EditarRegistro();
			ModificaGrilla(true);
			if (bNumera)
				txtCodigo.ReadOnly = true;
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtDescripcion, "");
			errorProvider.SetError(cmbFecha, "");
			errorProvider.SetError(txtCodigo, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (!bNumera)
				if (txtCodigo.Text.Trim().Length == 0)
				{
					bOk = false;
					errorProvider.SetError(txtCodigo, "Ingrese Codigo");
					txtCodigo.Select();
				}
				else
				{
					string stSelect = "Select Count(*) From Presupuesto "
						+ " Where Codigo = '" + txtCodigo.Text.Trim() + "'";
					int iCuenta = Funcion.iEscalarSQL(cdsPresupuesto, stSelect, false);
					if (!barraDatoSQL1.bNuevo)	iCuenta --;
					if (iCuenta > 0) 
					{
						bOk = false;
						errorProvider.SetError(txtCodigo, "Codigo ya existe");
						txtCodigo.Select();
					}					
				}

			if (txtDescripcion.Text.Trim().Length == 0)
			{
				bOk = false;
				errorProvider.SetError(txtDescripcion, "Ingrese Descripcion");
				txtDescripcion.Select();
			}

			if (cmbFecha.Value == System.DBNull.Value)
			{
				bOk = false;
				errorProvider.SetError(cmbFecha, "Ingrese Fecha");
				cmbFecha.Select();
			}
			else
			{
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				if (dtFecha.Year < 2000)
				{
					bOk = false;
					errorProvider.SetError(cmbFecha, "Fechas Validas desde el año 2000");
					cmbFecha.Select();
				}
			}

			return bOk;
		}
		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			if (barraDatoSQL1.bNuevo && bNumera) 
			{
				drPres["Codigo"] = Contabilidad.NumeroAsiento(cdsCuenta, 8);
			}
			barraDatoSQL1.GrabaRegistro();		
			ModificaGrilla(false);
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir Presuepuesto",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			RepPresupuesto miRep = new RepPresupuesto();
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		C1.Data.C1DataRow drPres;
		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			if (cdsPresupuesto.TableViews["Presupuesto"].Rows.Count > 0)
			drPres = cdsPresupuesto.TableViews["Presupuesto"].Rows[0];		
			btTotal_Click(this, e);
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		}

		private bool bNumera = true;
		private Acceso miAcceso;
		private void Presupuesto_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Presupuesto'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			miAcceso = new Acceso("0427", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			barraDatoSQL1.PosUltima();
			barraDatoSQL1.HabilitaControles(false);
			C1.Data.C1DataRow dr = this.cdsSeteoC.TableViews["SeteoC"].Rows[0];
//			if (dr["VerSucursal"] == System.DBNull.Value || !(bool) dr["VerSucursal"]) 
//				spnSucursal.Width = 0;
			bNumera = Funcion.bEscalarSQL(cdsCuenta, "SELECT Automatico" + " FROM AsientoTipo Where idTipoAsiento = 8", false);
			ModificaGrilla(false);
			Personalizaciones();
			LlenaCombos();
		}
		private void LlenaCombos()
		{		
			#region Proyecto Altropico
			string stExec = "Exec ProyectoCarga " + 1;
//			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'FCCCodigo'";
//			if (Funcion.bEjecutaSQL(cdsSeteoF, stActivado))
//			{
//				cmbProyecto.DisplayMember = "CentroCosto";
//			}
			cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);

			//cmbProyecto1.DataSource = cmbProyecto.DataSource;
			cmbProyecto.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;
			//cmbProyecto1.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;
			#endregion Proyecto Altropico
		}
		private void Personalizaciones()
		{
			#region Columna Proyecto*CC
			if (Funcion.bEjecutaSQL(cdsPresupuesto, "Exec SeteoGExiste 'FTRCCPA'")) 				
			{				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;										
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idSubProyecto"].Hidden = true;				
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["ProyectoBR"].Hidden = false;										
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["CentroCostoBr"].Hidden = false;	
			
			}
			#endregion Columna Proyecto*CC			
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			C1.Data.C1DataRow drSeteoC = this.cdsSeteoC.TableViews["SeteoC"].Rows[0];
			e.Layout.Bands[0].Columns["idProyecto"].Hidden = !(bool) drSeteoC["ProyectoLinea"];
			e.Layout.Bands[0].Columns["idSubProyecto"].Hidden = !(bool) drSeteoC["SubProyLinea"];
		}

		private void Presupuesto_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Presupuesto'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region codigo para filtrar cc por proy altropico
			if (Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'FTRCCPA'")) 				
			{
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "CentroCostoBR")
					{
						if(controlW != e.Cell.Row.Cells["CentroCostoBR"].Value.ToString())
						{
							try
							{
								controlCCP();
								controlW = e.Cell.Row.Cells["CentroCostoBR"].Value.ToString();							
								NuevoCodigo = codCC(controlW);
								NuevoCodigo=NuevoCodigo.Substring(0,5);	
								if (NuevoCodigo.Length >0)
								{
									string stSelect = "Select idProyecto From Proyecto Where CentroCosto like '"
										+ NuevoCodigo + "%'";
									int IdProy = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
									e.Cell.Row.Cells["idProyecto"].Value = IdProy;
								}
							}
							catch{};

						}
						//						else return;
					}
				}
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "ProyectoBR")
					{
						try
						{
							if(ultraGrid1.Rows.Count < 1 )return;
							string CodCentroCosto = ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value.ToString();							
							string NuevoCodigo1 = codCC(CodCentroCosto);
							//NuevoCodigo1=NuevoCodigo1.Substring(0,5);	

							if(controlW1 != e.Cell.Row.Cells["ProyectoBR"].Value.ToString())
							{
								controlCPP(NuevoCodigo1);							
								controlW1 = e.Cell.Row.Cells["ProyectoBR"].Value.ToString();
								if (controlW1.Trim().Length >0)
								{

									string stSelect1 = "Select idSubProyecto From subProyecto Where nombre like '"
										+ controlW1 + "%'";
									int IdProy1 = Funcion.iEscalarSQL(cdsSeteoF, stSelect1, true);
									e.Cell.Row.Cells["idSubProyecto"].Value = IdProy1;
								}

							}
						}
						catch{}
					}	
					//					else return;
					
				}
			}
		
			#endregion codigo para filtrar cc por proy altropico

			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "idCuentaPresup")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idCuenta"].Value = e.Cell.Row.Cells["idCuentaPresup"].Value;
					}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["idCuentaPresup"].Value = e.Row.Cells["idCuenta"].Value;
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


#region Columna Proyecto*CC

		private void controlCCP()
		{
			int iCuenta =0;
			string txtE = ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRCentroC '" + txtE + "',1";
					iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from Proyecto where idProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa, true);							
						controlW = CentroC;							
						ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
						using (BuscaCentroCosto miBusca = new BuscaCentroCosto(txtE))
						{
							if (DialogResult.OK == miBusca.ShowDialog())
							{
								try
								{
									int id = miBusca.idBusca;
									string stEmpresa1 = "Select Nombre from Proyecto where idProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa1, true);									
									controlW = CentroC1;									
									ultraGrid1.ActiveRow.Cells["CentroCostoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
		}
		private void controlCPP(string cod1)
		{
			int iCuenta =0;
			string txtE = ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRProyecto '" + txtE + "',1,'" + cod1 + "'";
					iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from SubProyecto where idSubProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa, true);							
						controlW1 = CentroC;							
						ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
					
						using (BuscaCentroCosto miBusca1 = new BuscaCentroCosto(txtE,cod1))
						{
							if (DialogResult.OK == miBusca1.ShowDialog())
							{
								try
								{
									int id = miBusca1.idBusca;
									string stEmpresa1 = "Select Nombre from SubProyecto where idSubProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsSeteoF, stEmpresa1, true);									
									controlW1 = CentroC1;									
									ultraGrid1.ActiveRow.Cells["ProyectoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
		}

		private string codCC(string cod)
		{
			//devuelve el codigo del centro de costo
			string stCodCC = "Select CentroCosto from Proyecto where nombre like '%"+cod+"%'";
			string IdGlobal = Funcion.sEscalarSQL(cdsSeteoF, stCodCC, true);	
			return IdGlobal;
		}

#endregion Columna Proyecto*CC

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			double dDebe = 0;
			double dHaber = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Debe"].Value != DBNull.Value) dDebe += (double) dr.Cells["Debe"].Value;
				if (dr.Cells["Haber"].Value != DBNull.Value) dHaber += (double) dr.Cells["Haber"].Value;
			}
			txtTotalD.Value = dDebe;
			txtTotalH.Value = dHaber;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

	}
}
