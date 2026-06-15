using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de SeteoR.
	/// </summary>
	public class SeteoR : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private C1.Data.C1DataSet cdsSeteoRol;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor6;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDistribuir;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor7;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor16;
		private Infragistics.Win.Misc.UltraLabel ultraLabel16;
		private Infragistics.Win.Misc.UltraLabel ultraLabel17;
		private Infragistics.Win.Misc.UltraLabel ultraLabel18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor17;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor19;
		private Infragistics.Win.Misc.UltraLabel ultraLabel19;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optListaPrecio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor20;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor21;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor22;
		private Infragistics.Win.Misc.UltraLabel ultraLabel20;
		private Infragistics.Win.Misc.UltraLabel ultraLabel21;
		private Infragistics.Win.Misc.UltraLabel ultraLabel22;
		private Infragistics.Win.Misc.UltraLabel ultraLabel23;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel24;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFondoRes;
		private Infragistics.Win.Misc.UltraLabel ultraLabel25;
		private Infragistics.Win.Misc.UltraLabel ultraLabel26;
		private Infragistics.Win.Misc.UltraLabel ultraLabel27;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel28;
		private DevExpress.XtraGrid.GridControl grdConceptos;
		private DevExpress.XtraGrid.Views.Grid.GridView grvConceptos;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
		private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDecimoIII;
		private Infragistics.Win.Misc.UltraButton btRubro;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAportable;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtValorRet;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVacaciones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFondoResMes;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAtrasos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtHoraNocturna;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtHoraDoble;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtHoraExtra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIngTotales;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBaseRetenida;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSueldo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUtilidades;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAporteIess;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDecimoIV;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRentaRetenida;
		private Infragistics.Win.Misc.UltraLabel ultraLabel29;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSobreSueldo;
		private System.ComponentModel.IContainer components;

		public SeteoR()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoRol = new C1.Data.C1DataSet();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDistribuir = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor6 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTextEditor22 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor21 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor20 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optListaPrecio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraTextEditor19 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel19 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor18 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor17 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel18 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel17 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel16 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor16 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor7 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtSobreSueldo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel29 = new Infragistics.Win.Misc.UltraLabel();
			this.grdConceptos = new DevExpress.XtraGrid.GridControl();
			this.grvConceptos = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
			this.ultraLabel28 = new Infragistics.Win.Misc.UltraLabel();
			this.txtAportable = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel27 = new Infragistics.Win.Misc.UltraLabel();
			this.txtValorRet = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel26 = new Infragistics.Win.Misc.UltraLabel();
			this.txtVacaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel25 = new Infragistics.Win.Misc.UltraLabel();
			this.txtFondoResMes = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel24 = new Infragistics.Win.Misc.UltraLabel();
			this.txtFondoRes = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel23 = new Infragistics.Win.Misc.UltraLabel();
			this.txtAtrasos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel22 = new Infragistics.Win.Misc.UltraLabel();
			this.txtHoraNocturna = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel21 = new Infragistics.Win.Misc.UltraLabel();
			this.txtHoraDoble = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel20 = new Infragistics.Win.Misc.UltraLabel();
			this.txtHoraExtra = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.txtIngTotales = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			this.txtBaseRetenida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDecimoIII = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.txtSueldo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.txtUtilidades = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.txtAporteIess = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.txtDecimoIV = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.txtRentaRetenida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btRubro = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor22)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optListaPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtSobreSueldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdConceptos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvConceptos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAportable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVacaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoResMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoRes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAtrasos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHoraNocturna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHoraDoble)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHoraExtra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngTotales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseRetenida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoIII)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSueldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidades)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAporteIess)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoIV)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRentaRetenida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.ultraCheckEditor3);
			this.ultraTabPageControl1.Controls.Add(this.ultraCheckEditor2);
			this.ultraTabPageControl1.Controls.Add(this.ultraCheckEditor1);
			this.ultraTabPageControl1.Controls.Add(this.chkDistribuir);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel6);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor5);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel5);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor4);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel4);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor3);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel3);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel2);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel1);
			this.ultraTabPageControl1.Controls.Add(this.ultraTextEditor6);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(492, 366);
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoRol, "SeteoRol.IngTotales"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(16, 241);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(328, 22);
			this.ultraCheckEditor3.TabIndex = 17;
			this.ultraCheckEditor3.Text = "Ingresos Totales o Ingresos que Aportan";
			// 
			// cdsSeteoRol
			// 
			this.cdsSeteoRol.BindingContextCtrl = this;
			this.cdsSeteoRol.DataLibrary = "LibPersonal";
			this.cdsSeteoRol.DataLibraryUrl = "";
			this.cdsSeteoRol.DataSetDef = "dsSeteoRol";
			this.cdsSeteoRol.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoRol.Name = "cdsSeteoRol";
			this.cdsSeteoRol.SchemaClassName = "LibPersonal.DataClass";
			this.cdsSeteoRol.SchemaDef = null;
			this.cdsSeteoRol.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoRol_BeforeFill);
			// 
			// ultraCheckEditor2
			// 
			this.ultraCheckEditor2.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoRol, "SeteoRol.FiltroSucursal"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(16, 207);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(200, 21);
			this.ultraCheckEditor2.TabIndex = 16;
			this.ultraCheckEditor2.Text = "Filtro de sucursal por usuario";
			// 
			// ultraCheckEditor1
			// 
			this.ultraCheckEditor1.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoRol, "SeteoRol.ChequeTodos"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(16, 172);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(392, 22);
			this.ultraCheckEditor1.TabIndex = 15;
			this.ultraCheckEditor1.Text = "Cheques para todos / Solo los que no van al cash management";
			// 
			// chkDistribuir
			// 
			this.chkDistribuir.BackColor = System.Drawing.Color.Transparent;
			this.chkDistribuir.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.cdsSeteoRol, "SeteoRol.Distribuir"));
			this.chkDistribuir.Location = new System.Drawing.Point(16, 138);
			this.chkDistribuir.Name = "chkDistribuir";
			this.chkDistribuir.Size = new System.Drawing.Size(272, 21);
			this.chkDistribuir.TabIndex = 14;
			this.chkDistribuir.Text = "Distribuir Asiento por Porcentajes de Distribución";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.SecPatronal"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(128, 17);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(88, 22);
			this.ultraTextEditor1.TabIndex = 0;
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel6.Location = new System.Drawing.Point(240, 86);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(52, 15);
			this.ultraLabel6.TabIndex = 11;
			this.ultraLabel6.Text = "Parroquia";
			// 
			// ultraTextEditor5
			// 
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Canton"));
			this.ultraTextEditor5.Location = new System.Drawing.Point(352, 52);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(88, 22);
			this.ultraTextEditor5.TabIndex = 8;
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel5.Location = new System.Drawing.Point(240, 52);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(39, 15);
			this.ultraLabel5.TabIndex = 9;
			this.ultraLabel5.Text = "Canton";
			// 
			// ultraTextEditor4
			// 
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Provincia"));
			this.ultraTextEditor4.Location = new System.Drawing.Point(352, 17);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(88, 22);
			this.ultraTextEditor4.TabIndex = 6;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Location = new System.Drawing.Point(240, 17);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(49, 15);
			this.ultraLabel4.TabIndex = 7;
			this.ultraLabel4.Text = "Provincia";
			// 
			// ultraTextEditor3
			// 
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.TipoSeguro"));
			this.ultraTextEditor3.Location = new System.Drawing.Point(128, 86);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(88, 22);
			this.ultraTextEditor3.TabIndex = 4;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(16, 86);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(64, 15);
			this.ultraLabel3.TabIndex = 5;
			this.ultraLabel3.Text = "Tipo Seguro";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.NumPatronal"));
			this.ultraTextEditor2.Location = new System.Drawing.Point(128, 52);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(88, 22);
			this.ultraTextEditor2.TabIndex = 2;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Location = new System.Drawing.Point(16, 52);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(88, 15);
			this.ultraLabel2.TabIndex = 3;
			this.ultraLabel2.Text = "Número Patronal";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(87, 15);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Sección Patronal";
			// 
			// ultraTextEditor6
			// 
			this.ultraTextEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Parroquia"));
			this.ultraTextEditor6.Location = new System.Drawing.Point(352, 86);
			this.ultraTextEditor6.Name = "ultraTextEditor6";
			this.ultraTextEditor6.Size = new System.Drawing.Size(88, 22);
			this.ultraTextEditor6.TabIndex = 10;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor22);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor21);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor20);
			this.ultraTabPageControl2.Controls.Add(this.optListaPrecio);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor19);
			this.ultraTabPageControl2.Controls.Add(this.ultraLabel19);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor18);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor17);
			this.ultraTabPageControl2.Controls.Add(this.ultraLabel18);
			this.ultraTabPageControl2.Controls.Add(this.ultraLabel17);
			this.ultraTabPageControl2.Controls.Add(this.ultraLabel16);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor16);
			this.ultraTabPageControl2.Controls.Add(this.ultraLabel7);
			this.ultraTabPageControl2.Controls.Add(this.ultraTextEditor7);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(492, 366);
			// 
			// ultraTextEditor22
			// 
			this.ultraTextEditor22.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojCodigo2"));
			this.ultraTextEditor22.Location = new System.Drawing.Point(304, 52);
			this.ultraTextEditor22.Name = "ultraTextEditor22";
			this.ultraTextEditor22.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor22.TabIndex = 28;
			// 
			// ultraTextEditor21
			// 
			this.ultraTextEditor21.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojTabla2"));
			this.ultraTextEditor21.Location = new System.Drawing.Point(304, 155);
			this.ultraTextEditor21.Name = "ultraTextEditor21";
			this.ultraTextEditor21.Size = new System.Drawing.Size(144, 22);
			this.ultraTextEditor21.TabIndex = 27;
			// 
			// ultraTextEditor20
			// 
			this.ultraTextEditor20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojTextoIO"));
			this.ultraTextEditor20.Location = new System.Drawing.Point(240, 250);
			this.ultraTextEditor20.Name = "ultraTextEditor20";
			this.ultraTextEditor20.Size = new System.Drawing.Size(32, 22);
			this.ultraTextEditor20.TabIndex = 26;
			this.toolTip1.SetToolTip(this.ultraTextEditor20, "Texto para Ingreso");
			// 
			// optListaPrecio
			// 
			this.optListaPrecio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojTipoIO"));
			this.optListaPrecio.ItemAppearance = appearance1;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Bit";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Número";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Texto";
			this.optListaPrecio.Items.Add(valueListItem1);
			this.optListaPrecio.Items.Add(valueListItem2);
			this.optListaPrecio.Items.Add(valueListItem3);
			this.optListaPrecio.ItemSpacingHorizontal = 10;
			this.optListaPrecio.ItemSpacingVertical = 7;
			this.optListaPrecio.Location = new System.Drawing.Point(120, 198);
			this.optListaPrecio.Name = "optListaPrecio";
			this.optListaPrecio.Size = new System.Drawing.Size(112, 78);
			this.optListaPrecio.TabIndex = 25;
			// 
			// ultraTextEditor19
			// 
			this.ultraTextEditor19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojTabla"));
			this.ultraTextEditor19.Location = new System.Drawing.Point(120, 155);
			this.ultraTextEditor19.Name = "ultraTextEditor19";
			this.ultraTextEditor19.Size = new System.Drawing.Size(176, 22);
			this.ultraTextEditor19.TabIndex = 24;
			// 
			// ultraLabel19
			// 
			this.ultraLabel19.AutoSize = true;
			this.ultraLabel19.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel19.Location = new System.Drawing.Point(24, 155);
			this.ultraLabel19.Name = "ultraLabel19";
			this.ultraLabel19.Size = new System.Drawing.Size(32, 15);
			this.ultraLabel19.TabIndex = 23;
			this.ultraLabel19.Text = "Tabla";
			// 
			// ultraTextEditor18
			// 
			this.ultraTextEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojEntrada"));
			this.ultraTextEditor18.Location = new System.Drawing.Point(120, 121);
			this.ultraTextEditor18.Name = "ultraTextEditor18";
			this.ultraTextEditor18.Size = new System.Drawing.Size(216, 22);
			this.ultraTextEditor18.TabIndex = 22;
			// 
			// ultraTextEditor17
			// 
			this.ultraTextEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojHora"));
			this.ultraTextEditor17.Location = new System.Drawing.Point(120, 86);
			this.ultraTextEditor17.Name = "ultraTextEditor17";
			this.ultraTextEditor17.Size = new System.Drawing.Size(216, 22);
			this.ultraTextEditor17.TabIndex = 21;
			// 
			// ultraLabel18
			// 
			this.ultraLabel18.AutoSize = true;
			this.ultraLabel18.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel18.Location = new System.Drawing.Point(24, 121);
			this.ultraLabel18.Name = "ultraLabel18";
			this.ultraLabel18.Size = new System.Drawing.Size(76, 15);
			this.ultraLabel18.TabIndex = 20;
			this.ultraLabel18.Text = "Entrada Salida";
			// 
			// ultraLabel17
			// 
			this.ultraLabel17.AutoSize = true;
			this.ultraLabel17.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel17.Location = new System.Drawing.Point(24, 86);
			this.ultraLabel17.Name = "ultraLabel17";
			this.ultraLabel17.Size = new System.Drawing.Size(41, 15);
			this.ultraLabel17.TabIndex = 19;
			this.ultraLabel17.Text = "Tiempo";
			// 
			// ultraLabel16
			// 
			this.ultraLabel16.AutoSize = true;
			this.ultraLabel16.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel16.Location = new System.Drawing.Point(24, 52);
			this.ultraLabel16.Name = "ultraLabel16";
			this.ultraLabel16.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel16.TabIndex = 18;
			this.ultraLabel16.Text = "Código";
			// 
			// ultraTextEditor16
			// 
			this.ultraTextEditor16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojCodigo"));
			this.ultraTextEditor16.Location = new System.Drawing.Point(120, 52);
			this.ultraTextEditor16.Name = "ultraTextEditor16";
			this.ultraTextEditor16.Size = new System.Drawing.Size(176, 22);
			this.ultraTextEditor16.TabIndex = 17;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel7.Location = new System.Drawing.Point(24, 17);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(85, 15);
			this.ultraLabel7.TabIndex = 16;
			this.ultraLabel7.Text = "Reloj Biométrico";
			// 
			// ultraTextEditor7
			// 
			this.ultraTextEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RelojBiometrico"));
			this.ultraTextEditor7.Location = new System.Drawing.Point(120, 17);
			this.ultraTextEditor7.Name = "ultraTextEditor7";
			this.ultraTextEditor7.Size = new System.Drawing.Size(216, 22);
			this.ultraTextEditor7.TabIndex = 15;
			this.toolTip1.SetToolTip(this.ultraTextEditor7, "Ruta a archivo mdb");
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.txtSobreSueldo);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel29);
			this.ultraTabPageControl3.Controls.Add(this.grdConceptos);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel28);
			this.ultraTabPageControl3.Controls.Add(this.txtAportable);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel27);
			this.ultraTabPageControl3.Controls.Add(this.txtValorRet);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel26);
			this.ultraTabPageControl3.Controls.Add(this.txtVacaciones);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel25);
			this.ultraTabPageControl3.Controls.Add(this.txtFondoResMes);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel24);
			this.ultraTabPageControl3.Controls.Add(this.txtFondoRes);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel23);
			this.ultraTabPageControl3.Controls.Add(this.txtAtrasos);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel22);
			this.ultraTabPageControl3.Controls.Add(this.txtHoraNocturna);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel21);
			this.ultraTabPageControl3.Controls.Add(this.txtHoraDoble);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel20);
			this.ultraTabPageControl3.Controls.Add(this.txtHoraExtra);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel14);
			this.ultraTabPageControl3.Controls.Add(this.txtIngTotales);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel15);
			this.ultraTabPageControl3.Controls.Add(this.txtBaseRetenida);
			this.ultraTabPageControl3.Controls.Add(this.txtDecimoIII);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel8);
			this.ultraTabPageControl3.Controls.Add(this.txtSueldo);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel9);
			this.ultraTabPageControl3.Controls.Add(this.txtUtilidades);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel10);
			this.ultraTabPageControl3.Controls.Add(this.txtAporteIess);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel11);
			this.ultraTabPageControl3.Controls.Add(this.txtDecimoIV);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel12);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel13);
			this.ultraTabPageControl3.Controls.Add(this.txtRentaRetenida);
			this.ultraTabPageControl3.Controls.Add(this.btRubro);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(492, 366);
			// 
			// txtSobreSueldo
			// 
			this.txtSobreSueldo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.SobreSueldo"));
			this.txtSobreSueldo.Location = new System.Drawing.Point(120, 16);
			this.txtSobreSueldo.Name = "txtSobreSueldo";
			this.txtSobreSueldo.Size = new System.Drawing.Size(312, 22);
			this.txtSobreSueldo.TabIndex = 47;
			// 
			// ultraLabel29
			// 
			this.ultraLabel29.AutoSize = true;
			this.ultraLabel29.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel29.Location = new System.Drawing.Point(24, 16);
			this.ultraLabel29.Name = "ultraLabel29";
			this.ultraLabel29.Size = new System.Drawing.Size(70, 15);
			this.ultraLabel29.TabIndex = 48;
			this.ultraLabel29.Text = "Sobre Sueldo";
			// 
			// grdConceptos
			// 
			this.grdConceptos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdConceptos.EmbeddedNavigator
			// 
			this.grdConceptos.EmbeddedNavigator.Name = "";
			this.grdConceptos.Location = new System.Drawing.Point(232, 152);
			this.grdConceptos.MainView = this.grvConceptos;
			this.grdConceptos.Name = "grdConceptos";
			this.grdConceptos.Size = new System.Drawing.Size(248, 177);
			this.grdConceptos.TabIndex = 46;
			this.grdConceptos.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.grvConceptos});
			// 
			// grvConceptos
			// 
			this.grvConceptos.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																												this.gridColumn1,
																																												this.gridColumn2});
			this.grvConceptos.GridControl = this.grdConceptos;
			this.grvConceptos.Name = "grvConceptos";
			this.grvConceptos.OptionsView.ShowAutoFilterRow = true;
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
			this.gridColumn1.Visible = true;
			this.gridColumn1.VisibleIndex = 0;
			this.gridColumn1.Width = 137;
			// 
			// gridColumn2
			// 
			this.gridColumn2.Caption = "Valor";
			this.gridColumn2.FieldName = "Valor";
			this.gridColumn2.Name = "gridColumn2";
			this.gridColumn2.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
			this.gridColumn2.OptionsColumn.ReadOnly = true;
			this.gridColumn2.Visible = true;
			this.gridColumn2.VisibleIndex = 1;
			this.gridColumn2.Width = 90;
			// 
			// ultraLabel28
			// 
			this.ultraLabel28.AutoSize = true;
			this.ultraLabel28.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel28.Location = new System.Drawing.Point(24, 280);
			this.ultraLabel28.Name = "ultraLabel28";
			this.ultraLabel28.Size = new System.Drawing.Size(52, 15);
			this.ultraLabel28.TabIndex = 45;
			this.ultraLabel28.Text = "Aportable";
			// 
			// txtAportable
			// 
			this.txtAportable.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Aportable"));
			this.txtAportable.Location = new System.Drawing.Point(120, 280);
			this.txtAportable.Name = "txtAportable";
			this.txtAportable.Size = new System.Drawing.Size(88, 22);
			this.txtAportable.TabIndex = 44;
			this.txtAportable.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtAportable_MouseUp);
			// 
			// ultraLabel27
			// 
			this.ultraLabel27.AutoSize = true;
			this.ultraLabel27.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel27.Location = new System.Drawing.Point(24, 256);
			this.ultraLabel27.Name = "ultraLabel27";
			this.ultraLabel27.Size = new System.Drawing.Size(69, 15);
			this.ultraLabel27.TabIndex = 43;
			this.ultraLabel27.Text = "I.R. Asumido";
			// 
			// txtValorRet
			// 
			this.txtValorRet.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.ValRet"));
			this.txtValorRet.Location = new System.Drawing.Point(120, 256);
			this.txtValorRet.Name = "txtValorRet";
			this.txtValorRet.Size = new System.Drawing.Size(88, 22);
			this.txtValorRet.TabIndex = 42;
			this.txtValorRet.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtValorRet_MouseUp);
			// 
			// ultraLabel26
			// 
			this.ultraLabel26.AutoSize = true;
			this.ultraLabel26.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel26.Location = new System.Drawing.Point(24, 328);
			this.ultraLabel26.Name = "ultraLabel26";
			this.ultraLabel26.Size = new System.Drawing.Size(59, 15);
			this.ultraLabel26.TabIndex = 41;
			this.ultraLabel26.Text = "Vacaciones";
			// 
			// txtVacaciones
			// 
			this.txtVacaciones.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Vacacion"));
			this.txtVacaciones.Location = new System.Drawing.Point(120, 328);
			this.txtVacaciones.Name = "txtVacaciones";
			this.txtVacaciones.Size = new System.Drawing.Size(88, 22);
			this.txtVacaciones.TabIndex = 40;
			this.txtVacaciones.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtVacaciones_MouseUp);
			// 
			// ultraLabel25
			// 
			this.ultraLabel25.AutoSize = true;
			this.ultraLabel25.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel25.Location = new System.Drawing.Point(24, 112);
			this.ultraLabel25.Name = "ultraLabel25";
			this.ultraLabel25.Size = new System.Drawing.Size(89, 15);
			this.ultraLabel25.TabIndex = 39;
			this.ultraLabel25.Text = "Fondo Resv. Mes";
			// 
			// txtFondoResMes
			// 
			this.txtFondoResMes.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.FondoResMes"));
			this.txtFondoResMes.Location = new System.Drawing.Point(120, 112);
			this.txtFondoResMes.Name = "txtFondoResMes";
			this.txtFondoResMes.Size = new System.Drawing.Size(88, 22);
			this.txtFondoResMes.TabIndex = 38;
			this.txtFondoResMes.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtFondoResMes_MouseUp);
			// 
			// ultraLabel24
			// 
			this.ultraLabel24.AutoSize = true;
			this.ultraLabel24.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel24.Location = new System.Drawing.Point(24, 136);
			this.ultraLabel24.Name = "ultraLabel24";
			this.ultraLabel24.Size = new System.Drawing.Size(66, 15);
			this.ultraLabel24.TabIndex = 37;
			this.ultraLabel24.Text = "Fondo Resv.";
			// 
			// txtFondoRes
			// 
			this.txtFondoRes.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.FondoRes"));
			this.txtFondoRes.Location = new System.Drawing.Point(120, 136);
			this.txtFondoRes.Name = "txtFondoRes";
			this.txtFondoRes.Size = new System.Drawing.Size(88, 22);
			this.txtFondoRes.TabIndex = 36;
			this.txtFondoRes.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtFondoRes_MouseUp);
			// 
			// ultraLabel23
			// 
			this.ultraLabel23.AutoSize = true;
			this.ultraLabel23.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel23.Location = new System.Drawing.Point(232, 40);
			this.ultraLabel23.Name = "ultraLabel23";
			this.ultraLabel23.Size = new System.Drawing.Size(36, 15);
			this.ultraLabel23.TabIndex = 35;
			this.ultraLabel23.Text = "Atraso";
			// 
			// txtAtrasos
			// 
			this.txtAtrasos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Atrazo"));
			this.txtAtrasos.Location = new System.Drawing.Point(344, 40);
			this.txtAtrasos.Name = "txtAtrasos";
			this.txtAtrasos.Size = new System.Drawing.Size(88, 22);
			this.txtAtrasos.TabIndex = 34;
			this.txtAtrasos.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtAtrasos_MouseUp);
			// 
			// ultraLabel22
			// 
			this.ultraLabel22.AutoSize = true;
			this.ultraLabel22.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel22.Location = new System.Drawing.Point(232, 112);
			this.ultraLabel22.Name = "ultraLabel22";
			this.ultraLabel22.Size = new System.Drawing.Size(64, 15);
			this.ultraLabel22.TabIndex = 33;
			this.ultraLabel22.Text = "H. Nocturna";
			// 
			// txtHoraNocturna
			// 
			this.txtHoraNocturna.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.HNocturna"));
			this.txtHoraNocturna.Location = new System.Drawing.Point(344, 112);
			this.txtHoraNocturna.Name = "txtHoraNocturna";
			this.txtHoraNocturna.Size = new System.Drawing.Size(88, 22);
			this.txtHoraNocturna.TabIndex = 32;
			this.txtHoraNocturna.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtHoraNocturna_MouseUp);
			// 
			// ultraLabel21
			// 
			this.ultraLabel21.AutoSize = true;
			this.ultraLabel21.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel21.Location = new System.Drawing.Point(232, 88);
			this.ultraLabel21.Name = "ultraLabel21";
			this.ultraLabel21.Size = new System.Drawing.Size(47, 15);
			this.ultraLabel21.TabIndex = 31;
			this.ultraLabel21.Text = "H. Doble";
			// 
			// txtHoraDoble
			// 
			this.txtHoraDoble.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.HDoble"));
			this.txtHoraDoble.Location = new System.Drawing.Point(344, 88);
			this.txtHoraDoble.Name = "txtHoraDoble";
			this.txtHoraDoble.Size = new System.Drawing.Size(88, 22);
			this.txtHoraDoble.TabIndex = 30;
			this.txtHoraDoble.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtHoraDoble_MouseUp);
			// 
			// ultraLabel20
			// 
			this.ultraLabel20.AutoSize = true;
			this.ultraLabel20.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel20.Location = new System.Drawing.Point(232, 64);
			this.ultraLabel20.Name = "ultraLabel20";
			this.ultraLabel20.Size = new System.Drawing.Size(45, 15);
			this.ultraLabel20.TabIndex = 29;
			this.ultraLabel20.Text = "H. Extra";
			// 
			// txtHoraExtra
			// 
			this.txtHoraExtra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.HExtra"));
			this.txtHoraExtra.Location = new System.Drawing.Point(344, 64);
			this.txtHoraExtra.Name = "txtHoraExtra";
			this.txtHoraExtra.Size = new System.Drawing.Size(88, 22);
			this.txtHoraExtra.TabIndex = 28;
			this.txtHoraExtra.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtHoraExtra_MouseUp);
			// 
			// ultraLabel14
			// 
			this.ultraLabel14.AutoSize = true;
			this.ultraLabel14.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel14.Location = new System.Drawing.Point(24, 208);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(75, 15);
			this.ultraLabel14.TabIndex = 27;
			this.ultraLabel14.Text = "Base Retenida";
			// 
			// txtIngTotales
			// 
			this.txtIngTotales.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Comisiones"));
			this.txtIngTotales.Location = new System.Drawing.Point(120, 304);
			this.txtIngTotales.Name = "txtIngTotales";
			this.txtIngTotales.Size = new System.Drawing.Size(88, 22);
			this.txtIngTotales.TabIndex = 24;
			this.txtIngTotales.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtIngTotales_MouseUp);
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.AutoSize = true;
			this.ultraLabel15.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel15.Location = new System.Drawing.Point(24, 304);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(87, 15);
			this.ultraLabel15.TabIndex = 25;
			this.ultraLabel15.Text = "Ingresos Totales";
			// 
			// txtBaseRetenida
			// 
			this.txtBaseRetenida.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.BaseRenta"));
			this.txtBaseRetenida.Location = new System.Drawing.Point(120, 208);
			this.txtBaseRetenida.Name = "txtBaseRetenida";
			this.txtBaseRetenida.Size = new System.Drawing.Size(88, 22);
			this.txtBaseRetenida.TabIndex = 26;
			this.txtBaseRetenida.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtBaseRetenida_MouseUp);
			// 
			// txtDecimoIII
			// 
			this.txtDecimoIII.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.DecimoXIII"));
			this.txtDecimoIII.Location = new System.Drawing.Point(120, 64);
			this.txtDecimoIII.Name = "txtDecimoIII";
			this.txtDecimoIII.Size = new System.Drawing.Size(88, 22);
			this.txtDecimoIII.TabIndex = 12;
			this.txtDecimoIII.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtDecimoIII_MouseUp);
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel8.Location = new System.Drawing.Point(24, 232);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(81, 15);
			this.ultraLabel8.TabIndex = 23;
			this.ultraLabel8.Text = "Renta Retenida";
			// 
			// txtSueldo
			// 
			this.txtSueldo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Sueldo"));
			this.txtSueldo.Location = new System.Drawing.Point(120, 40);
			this.txtSueldo.Name = "txtSueldo";
			this.txtSueldo.Size = new System.Drawing.Size(88, 22);
			this.txtSueldo.TabIndex = 20;
			this.txtSueldo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtSueldo_MouseUp);
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.AutoSize = true;
			this.ultraLabel9.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel9.Location = new System.Drawing.Point(24, 40);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel9.TabIndex = 21;
			this.ultraLabel9.Text = "Sueldo";
			// 
			// txtUtilidades
			// 
			this.txtUtilidades.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.Utilidades"));
			this.txtUtilidades.Location = new System.Drawing.Point(120, 160);
			this.txtUtilidades.Name = "txtUtilidades";
			this.txtUtilidades.Size = new System.Drawing.Size(88, 22);
			this.txtUtilidades.TabIndex = 18;
			this.txtUtilidades.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtUtilidades_MouseUp);
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.AutoSize = true;
			this.ultraLabel10.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel10.Location = new System.Drawing.Point(24, 160);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(53, 15);
			this.ultraLabel10.TabIndex = 19;
			this.ultraLabel10.Text = "Utilidades";
			// 
			// txtAporteIess
			// 
			this.txtAporteIess.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.AporteIess"));
			this.txtAporteIess.Location = new System.Drawing.Point(120, 184);
			this.txtAporteIess.Name = "txtAporteIess";
			this.txtAporteIess.Size = new System.Drawing.Size(88, 22);
			this.txtAporteIess.TabIndex = 16;
			this.txtAporteIess.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtAporteIess_MouseUp);
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.AutoSize = true;
			this.ultraLabel11.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel11.Location = new System.Drawing.Point(24, 184);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(64, 15);
			this.ultraLabel11.TabIndex = 17;
			this.ultraLabel11.Text = "Aporte IESS";
			// 
			// txtDecimoIV
			// 
			this.txtDecimoIV.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.DecimoXIV"));
			this.txtDecimoIV.Location = new System.Drawing.Point(120, 88);
			this.txtDecimoIV.Name = "txtDecimoIV";
			this.txtDecimoIV.Size = new System.Drawing.Size(88, 22);
			this.txtDecimoIV.TabIndex = 14;
			this.txtDecimoIV.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtDecimoIV_MouseUp);
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.AutoSize = true;
			this.ultraLabel12.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel12.Location = new System.Drawing.Point(24, 88);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(62, 15);
			this.ultraLabel12.TabIndex = 15;
			this.ultraLabel12.Text = "Décimo XIV";
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.AutoSize = true;
			this.ultraLabel13.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel13.Location = new System.Drawing.Point(24, 64);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(64, 15);
			this.ultraLabel13.TabIndex = 13;
			this.ultraLabel13.Text = "Décimo XIII";
			// 
			// txtRentaRetenida
			// 
			this.txtRentaRetenida.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRol, "SeteoRol.RentaRetenida"));
			this.txtRentaRetenida.Location = new System.Drawing.Point(120, 232);
			this.txtRentaRetenida.Name = "txtRentaRetenida";
			this.txtRentaRetenida.Size = new System.Drawing.Size(88, 22);
			this.txtRentaRetenida.TabIndex = 22;
			this.txtRentaRetenida.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtRentaRetenida_MouseUp);
			// 
			// btRubro
			// 
			this.btRubro.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btRubro.Location = new System.Drawing.Point(232, 328);
			this.btRubro.Name = "btRubro";
			this.btRubro.Size = new System.Drawing.Size(75, 25);
			this.btRubro.TabIndex = 18;
			this.btRubro.Text = "&Rubro";
			this.btRubro.Click += new System.EventHandler(this.btRubro_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(64, 408);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 12;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(344, 408);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 13;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 8);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(496, 392);
			this.ultraTabControl1.TabIndex = 17;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Empresa";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Reloj";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Campos";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(492, 366);
			// 
			// SeteoR
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(528, 446);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "SeteoR";
			this.Text = "Seteo de Roles de Pago";
			this.Load += new System.EventHandler(this.SeteoR_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor6)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor22)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optListaPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor7)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtSobreSueldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdConceptos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvConceptos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAportable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVacaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoResMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoRes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAtrasos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHoraNocturna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHoraDoble)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHoraExtra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngTotales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseRetenida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoIII)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSueldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidades)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAporteIess)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoIV)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRentaRetenida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void SeteoR_Load(object sender, System.EventArgs e)
		{
			grdConceptos.DataSource = Funcion.dvProcedimiento(cdsSeteoRol, 
				"Select distinct(rc.Concepto) As Codigos, (RTrim(rt.Tipo) + cast(rc.Posicion as varchar)) As Valor "+
				"From RolConcepto rc Inner Join RolTipo rt On rc.idRolTipo = rt.idRolTipo Order by rc.Concepto");
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsSeteoRol.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Grabar Rol");
				return;
			}
			Close();
		}

		private void cdsSeteoRol_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoRol.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void btRubro_Click(object sender, System.EventArgs e)
		{
			int idConcepto = 0;
			int idRolGrupo = 0;
			if (grvConceptos.FocusedRowHandle < 0)
			{
				MessageBox.Show("Seleccione columna que desea ver el Rubro", "Información");
				return;
			}
			string stColumna = grvConceptos.GetRowCellValue(grvConceptos.FocusedRowHandle, "Valor").ToString().ToLower();
			BuscaRubro(stColumna);
		}

		private void BuscaRubro(string stColumna)
		{
			stColumna = stColumna.ToLower();
			int idTipo = 4;
			int idConcepto = 0;
			int idRolGrupo = 0;

			string stTipo = "valor";
			if (stColumna.StartsWith("ingreso")) {idTipo = 1; stTipo = "ingreso";}
			else if (stColumna.StartsWith("egreso")) {idTipo = 2; stTipo = "egreso";}
			else if (stColumna.StartsWith("provision")) {idTipo = 3; stTipo = "provision";}
			else if (stColumna.StartsWith("valor")) {idTipo = 4; stTipo = "valor";}
			else
			{
				MessageBox.Show("Columna no contiene un Rubro", "Información");
				return;
			}
			int iPos = stColumna.IndexOf(stTipo);
			stTipo = stColumna.Substring(stTipo.Length);
			iPos = int.Parse(stTipo);
			string stSelect = string.Format("Select Count(*) From RolConcepto Where idRolTipo = {0} And Posicion = {1} And (idRolGrupo = {2} Or {2} = 0)",
				idTipo, iPos, idRolGrupo);

			idConcepto = Funcion.iEscalarSQL(cdsSeteoRol, stSelect);
			if (idConcepto == 0)
			{
				MessageBox.Show("Rubro no existe en grupo seleccionado", "Información");
				return;
			}
			stSelect = stSelect.Replace("Count(*)", "Top 1 idRolConcepto");
			idConcepto = Funcion.iEscalarSQL(cdsSeteoRol, stSelect);
			RolConcepto miConcepto = new RolConcepto(idConcepto);
			miConcepto.MdiParent = MdiParent;
			miConcepto.Show();
		}

		private void txtDecimoIII_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtDecimoIII.Text);
		}

		private void txtDecimoIV_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtDecimoIV.Text);
		}

		private void txtAporteIess_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtAporteIess.Text);
		}

		private void txtUtilidades_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtUtilidades.Text);
		}

		private void txtSueldo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtSueldo.Text);
		}

		private void txtRentaRetenida_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtRentaRetenida.Text);
		}

		private void txtFondoRes_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtFondoRes.Text);
		}

		private void txtVacaciones_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtVacaciones.Text);
		}

		private void txtAportable_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtAportable.Text);
		}

		private void txtAtrasos_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtAtrasos.Text);
		}

		private void txtFondoResMes_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtFondoResMes.Text);
		}

		private void txtValorRet_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtValorRet.Text);
		}

		private void txtIngTotales_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtIngTotales.Text);
		}

		private void txtBaseRetenida_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtBaseRetenida.Text);
		}

		private void txtHoraExtra_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtHoraExtra.Text);
		}

		private void txtHoraDoble_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtHoraDoble.Text);
		}

		private void txtHoraNocturna_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			BuscaRubro(txtHoraNocturna.Text);
		}
	}
}
