using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de SeteoC.
	/// </summary>
	public class SeteoC : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private C1.Data.C1DataSet cdsSeteoC;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private System.Windows.Forms.Label label8;
		private DevExpress.XtraEditors.TextEdit txtCapital;
		private DevExpress.XtraEditors.SimpleButton btPlanCuenta;
		private DevExpress.XtraEditors.CheckEdit chkSubCentro;
		private DevExpress.XtraEditors.CheckEdit chkCentroCosto;
		private DevExpress.XtraEditors.CheckEdit chkCopiaBenef;
		private DevExpress.XtraEditors.CheckEdit chkCopiaNota;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.TextEdit txtEgresos;
		private DevExpress.XtraEditors.TextEdit txtIngresos;
		private DevExpress.XtraEditors.TextEdit txtPasivo;
		private DevExpress.XtraEditors.LookUpEdit cmbCuentaU;
		private DevExpress.XtraEditors.DateEdit dtFechaFin;
		private DevExpress.XtraEditors.DateEdit dtFechaIni;
		private DevExpress.XtraEditors.TextEdit txtActivo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsTipoAsiento;
		private DevExpress.XtraEditors.CheckEdit checkEdit1;
		private DevExpress.XtraEditors.CheckEdit checkEdit2;
		private DevExpress.XtraEditors.CheckEdit checkEdit3;
		private DevExpress.XtraEditors.CheckEdit checkEdit4;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGastoAT_F;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private DevExpress.XtraEditors.CheckEdit checkEdit5;
		private DevExpress.XtraEditors.CheckEdit checkEdit6;
		private DevExpress.XtraEditors.CheckEdit checkEdit7;
		private DevExpress.XtraEditors.CheckEdit checkEdit8;
		private DevExpress.XtraEditors.CheckEdit checkEdit9;
		private DevExpress.XtraEditors.SimpleButton btListaCuenta;
		private DevExpress.XtraEditors.CheckEdit checkEdit10;
		private DevExpress.XtraEditors.CheckEdit checkEdit11;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public SeteoC()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AsientoTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btListaCuenta = new DevExpress.XtraEditors.SimpleButton();
			this.checkEdit9 = new DevExpress.XtraEditors.CheckEdit();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.checkEdit6 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit4 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit3 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit2 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit1 = new DevExpress.XtraEditors.CheckEdit();
			this.label8 = new System.Windows.Forms.Label();
			this.txtCapital = new DevExpress.XtraEditors.TextEdit();
			this.btPlanCuenta = new DevExpress.XtraEditors.SimpleButton();
			this.chkSubCentro = new DevExpress.XtraEditors.CheckEdit();
			this.chkCentroCosto = new DevExpress.XtraEditors.CheckEdit();
			this.chkCopiaBenef = new DevExpress.XtraEditors.CheckEdit();
			this.chkCopiaNota = new DevExpress.XtraEditors.CheckEdit();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtEgresos = new DevExpress.XtraEditors.TextEdit();
			this.txtIngresos = new DevExpress.XtraEditors.TextEdit();
			this.txtPasivo = new DevExpress.XtraEditors.TextEdit();
			this.cmbCuentaU = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.dtFechaFin = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaIni = new DevExpress.XtraEditors.DateEdit();
			this.txtActivo = new DevExpress.XtraEditors.TextEdit();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsTipoAsiento = new C1.Data.C1DataSet();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkGastoAT_F = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.checkEdit11 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit10 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit8 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit7 = new DevExpress.XtraEditors.CheckEdit();
			this.checkEdit5 = new DevExpress.XtraEditors.CheckEdit();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit9.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit6.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit4.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit3.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCapital.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentroCosto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCopiaBenef.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCopiaNota.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEgresos.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresos.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasivo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaU.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtActivo.Properties)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoAsiento)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit11.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit10.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit8.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit7.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit5.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btListaCuenta);
			this.ultraTabPageControl1.Controls.Add(this.checkEdit9);
			this.ultraTabPageControl1.Controls.Add(this.checkEdit6);
			this.ultraTabPageControl1.Controls.Add(this.checkEdit4);
			this.ultraTabPageControl1.Controls.Add(this.checkEdit3);
			this.ultraTabPageControl1.Controls.Add(this.checkEdit2);
			this.ultraTabPageControl1.Controls.Add(this.checkEdit1);
			this.ultraTabPageControl1.Controls.Add(this.label8);
			this.ultraTabPageControl1.Controls.Add(this.txtCapital);
			this.ultraTabPageControl1.Controls.Add(this.btPlanCuenta);
			this.ultraTabPageControl1.Controls.Add(this.chkSubCentro);
			this.ultraTabPageControl1.Controls.Add(this.chkCentroCosto);
			this.ultraTabPageControl1.Controls.Add(this.chkCopiaBenef);
			this.ultraTabPageControl1.Controls.Add(this.chkCopiaNota);
			this.ultraTabPageControl1.Controls.Add(this.label6);
			this.ultraTabPageControl1.Controls.Add(this.label7);
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Controls.Add(this.label4);
			this.ultraTabPageControl1.Controls.Add(this.label3);
			this.ultraTabPageControl1.Controls.Add(this.label2);
			this.ultraTabPageControl1.Controls.Add(this.label1);
			this.ultraTabPageControl1.Controls.Add(this.txtEgresos);
			this.ultraTabPageControl1.Controls.Add(this.txtIngresos);
			this.ultraTabPageControl1.Controls.Add(this.txtPasivo);
			this.ultraTabPageControl1.Controls.Add(this.cmbCuentaU);
			this.ultraTabPageControl1.Controls.Add(this.dtFechaFin);
			this.ultraTabPageControl1.Controls.Add(this.dtFechaIni);
			this.ultraTabPageControl1.Controls.Add(this.txtActivo);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(612, 370);
			// 
			// btListaCuenta
			// 
			this.btListaCuenta.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btListaCuenta.Location = new System.Drawing.Point(200, 327);
			this.btListaCuenta.Name = "btListaCuenta";
			this.btListaCuenta.Size = new System.Drawing.Size(24, 25);
			this.btListaCuenta.TabIndex = 77;
			this.btListaCuenta.Text = "...";
			this.btListaCuenta.Click += new System.EventHandler(this.btListaCuenta_Click);
			// 
			// checkEdit9
			// 
			this.checkEdit9.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.ListarCuentas"));
			this.checkEdit9.Location = new System.Drawing.Point(40, 327);
			this.checkEdit9.Name = "checkEdit9";
			// 
			// checkEdit9.Properties
			// 
			this.checkEdit9.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit9.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit9.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit9.Properties.Caption = "Ver Listas Desplegables";
			this.checkEdit9.Size = new System.Drawing.Size(152, 21);
			this.checkEdit9.TabIndex = 76;
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
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// checkEdit6
			// 
			this.checkEdit6.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.verFechaComp"));
			this.checkEdit6.Location = new System.Drawing.Point(320, 327);
			this.checkEdit6.Name = "checkEdit6";
			// 
			// checkEdit6.Properties
			// 
			this.checkEdit6.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit6.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit6.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit6.Properties.Caption = "Ver Fecha Comprobante";
			this.checkEdit6.Size = new System.Drawing.Size(256, 21);
			this.checkEdit6.TabIndex = 75;
			// 
			// checkEdit4
			// 
			this.checkEdit4.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.VerDeposito"));
			this.checkEdit4.Location = new System.Drawing.Point(320, 293);
			this.checkEdit4.Name = "checkEdit4";
			// 
			// checkEdit4.Properties
			// 
			this.checkEdit4.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit4.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit4.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit4.Properties.Caption = "Ver Depósito";
			this.checkEdit4.Size = new System.Drawing.Size(256, 21);
			this.checkEdit4.TabIndex = 74;
			// 
			// checkEdit3
			// 
			this.checkEdit3.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.BalanceProyecto"));
			this.checkEdit3.Location = new System.Drawing.Point(40, 293);
			this.checkEdit3.Name = "checkEdit3";
			// 
			// checkEdit3.Properties
			// 
			this.checkEdit3.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit3.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit3.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit3.Properties.Caption = "Balance por Proyecto";
			this.checkEdit3.Size = new System.Drawing.Size(192, 21);
			this.checkEdit3.TabIndex = 73;
			// 
			// checkEdit2
			// 
			this.checkEdit2.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.BalanceLinea"));
			this.checkEdit2.Location = new System.Drawing.Point(40, 258);
			this.checkEdit2.Name = "checkEdit2";
			// 
			// checkEdit2.Properties
			// 
			this.checkEdit2.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit2.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit2.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit2.Properties.Caption = "Balance por Centro de Costo";
			this.checkEdit2.Size = new System.Drawing.Size(192, 21);
			this.checkEdit2.TabIndex = 72;
			// 
			// checkEdit1
			// 
			this.checkEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.VerSucursal"));
			this.checkEdit1.Location = new System.Drawing.Point(320, 258);
			this.checkEdit1.Name = "checkEdit1";
			// 
			// checkEdit1.Properties
			// 
			this.checkEdit1.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit1.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit1.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit1.Properties.Caption = "Ver Sucursales";
			this.checkEdit1.Size = new System.Drawing.Size(256, 21);
			this.checkEdit1.TabIndex = 71;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(38, 155);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(38, 16);
			this.label8.TabIndex = 70;
			this.label8.Text = "Capital";
			// 
			// txtCapital
			// 
			this.txtCapital.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CuentaCapital"));
			this.txtCapital.EditValue = "";
			this.txtCapital.Location = new System.Drawing.Point(150, 155);
			this.txtCapital.Name = "txtCapital";
			// 
			// txtCapital.Properties
			// 
			this.txtCapital.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtCapital.Properties.Appearance.Options.UseBackColor = true;
			this.txtCapital.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtCapital.Size = new System.Drawing.Size(112, 20);
			this.txtCapital.TabIndex = 69;
			// 
			// btPlanCuenta
			// 
			this.btPlanCuenta.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btPlanCuenta.Location = new System.Drawing.Point(478, 52);
			this.btPlanCuenta.Name = "btPlanCuenta";
			this.btPlanCuenta.Size = new System.Drawing.Size(24, 24);
			this.btPlanCuenta.TabIndex = 68;
			this.btPlanCuenta.Text = "...";
			this.btPlanCuenta.Click += new System.EventHandler(this.btPlanCuenta_Click);
			// 
			// chkSubCentro
			// 
			this.chkSubCentro.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.SubProyLinea"));
			this.chkSubCentro.Location = new System.Drawing.Point(40, 224);
			this.chkSubCentro.Name = "chkSubCentro";
			// 
			// chkSubCentro.Properties
			// 
			this.chkSubCentro.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.chkSubCentro.Properties.Appearance.Options.UseBackColor = true;
			this.chkSubCentro.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkSubCentro.Properties.Caption = "Proyecto por línea";
			this.chkSubCentro.Size = new System.Drawing.Size(192, 21);
			this.chkSubCentro.TabIndex = 66;
			// 
			// chkCentroCosto
			// 
			this.chkCentroCosto.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.ProyectoLinea"));
			this.chkCentroCosto.Location = new System.Drawing.Point(40, 190);
			this.chkCentroCosto.Name = "chkCentroCosto";
			// 
			// chkCentroCosto.Properties
			// 
			this.chkCentroCosto.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.chkCentroCosto.Properties.Appearance.Options.UseBackColor = true;
			this.chkCentroCosto.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkCentroCosto.Properties.Caption = "Centro de Costo por línea";
			this.chkCentroCosto.Size = new System.Drawing.Size(192, 21);
			this.chkCentroCosto.TabIndex = 65;
			// 
			// chkCopiaBenef
			// 
			this.chkCopiaBenef.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CopiaBeneficiario"));
			this.chkCopiaBenef.Location = new System.Drawing.Point(320, 224);
			this.chkCopiaBenef.Name = "chkCopiaBenef";
			// 
			// chkCopiaBenef.Properties
			// 
			this.chkCopiaBenef.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.chkCopiaBenef.Properties.Appearance.Options.UseBackColor = true;
			this.chkCopiaBenef.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkCopiaBenef.Properties.Caption = "Copiar Beneficiario de Retención a Asiento";
			this.chkCopiaBenef.Size = new System.Drawing.Size(256, 21);
			this.chkCopiaBenef.TabIndex = 64;
			// 
			// chkCopiaNota
			// 
			this.chkCopiaNota.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CopiarNota"));
			this.chkCopiaNota.Location = new System.Drawing.Point(320, 190);
			this.chkCopiaNota.Name = "chkCopiaNota";
			// 
			// chkCopiaNota.Properties
			// 
			this.chkCopiaNota.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.chkCopiaNota.Properties.Appearance.Options.UseBackColor = true;
			this.chkCopiaNota.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkCopiaNota.Properties.Caption = "Copiar Nota de Retención a Asiento";
			this.chkCopiaNota.Size = new System.Drawing.Size(232, 21);
			this.chkCopiaNota.TabIndex = 63;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(326, 121);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(43, 16);
			this.label6.TabIndex = 62;
			this.label6.Text = "Egresos";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(326, 86);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(47, 16);
			this.label7.TabIndex = 60;
			this.label7.Text = "Ingresos";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(38, 121);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(36, 16);
			this.label5.TabIndex = 58;
			this.label5.Text = "Pasivo";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(38, 86);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 56;
			this.label4.Text = "Activo";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(38, 52);
			this.label3.Name = "label3";
			this.label3.TabIndex = 55;
			this.label3.Text = "Cuenta de Utilidad:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(326, 17);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(73, 16);
			this.label2.TabIndex = 52;
			this.label2.Text = "Periodo Final:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(40, 17);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(79, 16);
			this.label1.TabIndex = 50;
			this.label1.Text = "Periodo Inicial:";
			// 
			// txtEgresos
			// 
			this.txtEgresos.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CuentaEgreso"));
			this.txtEgresos.EditValue = "";
			this.txtEgresos.Location = new System.Drawing.Point(406, 121);
			this.txtEgresos.Name = "txtEgresos";
			// 
			// txtEgresos.Properties
			// 
			this.txtEgresos.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtEgresos.Properties.Appearance.Options.UseBackColor = true;
			this.txtEgresos.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtEgresos.Size = new System.Drawing.Size(96, 20);
			this.txtEgresos.TabIndex = 61;
			// 
			// txtIngresos
			// 
			this.txtIngresos.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CuentaIngreso"));
			this.txtIngresos.EditValue = "";
			this.txtIngresos.Location = new System.Drawing.Point(406, 86);
			this.txtIngresos.Name = "txtIngresos";
			// 
			// txtIngresos.Properties
			// 
			this.txtIngresos.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtIngresos.Properties.Appearance.Options.UseBackColor = true;
			this.txtIngresos.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIngresos.Size = new System.Drawing.Size(96, 20);
			this.txtIngresos.TabIndex = 59;
			// 
			// txtPasivo
			// 
			this.txtPasivo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CuentaPasivo"));
			this.txtPasivo.EditValue = "";
			this.txtPasivo.Location = new System.Drawing.Point(150, 121);
			this.txtPasivo.Name = "txtPasivo";
			// 
			// txtPasivo.Properties
			// 
			this.txtPasivo.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtPasivo.Properties.Appearance.Options.UseBackColor = true;
			this.txtPasivo.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtPasivo.Size = new System.Drawing.Size(112, 20);
			this.txtPasivo.TabIndex = 57;
			// 
			// cmbCuentaU
			// 
			this.cmbCuentaU.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CuentaUtilidad"));
			this.cmbCuentaU.Location = new System.Drawing.Point(150, 52);
			this.cmbCuentaU.Name = "cmbCuentaU";
			// 
			// cmbCuentaU.Properties
			// 
			this.cmbCuentaU.Properties.Appearance.BackColor = System.Drawing.SystemColors.Control;
			this.cmbCuentaU.Properties.Appearance.Options.UseBackColor = true;
			this.cmbCuentaU.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.cmbCuentaU.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCuentaU.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCuentaU.Properties.DataSource = this.cdvCuenta;
			this.cmbCuentaU.Properties.DisplayMember = "Codigo";
			this.cmbCuentaU.Properties.NullText = "";
			this.cmbCuentaU.Properties.ValueMember = "idCuenta";
			this.cmbCuentaU.Size = new System.Drawing.Size(320, 20);
			this.cmbCuentaU.TabIndex = 54;
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
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// dtFechaFin
			// 
			this.dtFechaFin.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaFin"));
			this.dtFechaFin.EditValue = new System.DateTime(2004, 9, 18, 0, 0, 0, 0);
			this.dtFechaFin.Location = new System.Drawing.Point(406, 17);
			this.dtFechaFin.Name = "dtFechaFin";
			// 
			// dtFechaFin.Properties
			// 
			this.dtFechaFin.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.dtFechaFin.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaFin.Size = new System.Drawing.Size(96, 20);
			this.dtFechaFin.TabIndex = 53;
			// 
			// dtFechaIni
			// 
			this.dtFechaIni.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaIni"));
			this.dtFechaIni.EditValue = new System.DateTime(2004, 9, 18, 0, 0, 0, 0);
			this.dtFechaIni.Location = new System.Drawing.Point(150, 17);
			this.dtFechaIni.Name = "dtFechaIni";
			// 
			// dtFechaIni.Properties
			// 
			this.dtFechaIni.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.dtFechaIni.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaIni.Size = new System.Drawing.Size(96, 20);
			this.dtFechaIni.TabIndex = 51;
			// 
			// txtActivo
			// 
			this.txtActivo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.CuentaActivo"));
			this.txtActivo.EditValue = "";
			this.txtActivo.Location = new System.Drawing.Point(150, 86);
			this.txtActivo.Name = "txtActivo";
			// 
			// txtActivo.Properties
			// 
			this.txtActivo.Properties.Appearance.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.txtActivo.Properties.Appearance.Options.UseBackColor = true;
			this.txtActivo.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtActivo.Size = new System.Drawing.Size(112, 20);
			this.txtActivo.TabIndex = 49;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10769);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(612, 370);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "AsientoTipo";
			this.ultraGrid1.DataSource = this.cdsTipoAsiento;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 8;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 6;
			ultraGridColumn4.Header.Caption = "Automático";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.Caption = "Código";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn6.Header.Caption = "Número";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
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
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(24, 17);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(568, 319);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// cdsTipoAsiento
			// 
			this.cdsTipoAsiento.BindingContextCtrl = this;
			this.cdsTipoAsiento.DataLibrary = "LibContabilidad";
			this.cdsTipoAsiento.DataLibraryUrl = "";
			this.cdsTipoAsiento.DataSetDef = "dsAsiento";
			this.cdsTipoAsiento.FillOnRequest = false;
			this.cdsTipoAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoAsiento.Name = "cdsTipoAsiento";
			this.cdsTipoAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsTipoAsiento.SchemaDef = null;
			this.cdsTipoAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoAsiento_BeforeFill);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.chkGastoAT_F);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel3);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor3);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel2);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel1);
			this.ultraTabPageControl3.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl3.Controls.Add(this.ultraCheckEditor1);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10769);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(612, 370);
			// 
			// chkGastoAT_F
			// 
			this.chkGastoAT_F.BackColor = System.Drawing.Color.Transparent;
			this.chkGastoAT_F.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoC, "SeteoC.GastoAT_F"));
			this.chkGastoAT_F.Location = new System.Drawing.Point(336, 26);
			this.chkGastoAT_F.Name = "chkGastoAT_F";
			this.chkGastoAT_F.Size = new System.Drawing.Size(224, 21);
			this.chkGastoAT_F.TabIndex = 7;
			this.chkGastoAT_F.Text = "Gastos Afectan Formularios y Anexos";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(40, 138);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(65, 15);
			this.ultraLabel3.TabIndex = 6;
			this.ultraLabel3.Text = "Institucional";
			// 
			// ultraTextEditor3
			// 
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoC, "SeteoC.Institucional"));
			this.ultraTextEditor3.Location = new System.Drawing.Point(112, 138);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(112, 22);
			this.ultraTextEditor3.TabIndex = 5;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel2.Location = new System.Drawing.Point(40, 103);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(58, 15);
			this.ultraLabel2.TabIndex = 4;
			this.ultraLabel2.Text = "Geográfico";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoC, "SeteoC.Geografico"));
			this.ultraTextEditor2.Location = new System.Drawing.Point(112, 103);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(112, 22);
			this.ultraTextEditor2.TabIndex = 3;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(40, 69);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(52, 15);
			this.ultraLabel1.TabIndex = 2;
			this.ultraLabel1.Text = "Funcional";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoC, "SeteoC.Funcional"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(112, 69);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(112, 22);
			this.ultraTextEditor1.TabIndex = 1;
			// 
			// ultraCheckEditor1
			// 
			this.ultraCheckEditor1.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoC, "SeteoC.Gobierno"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(48, 26);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(184, 21);
			this.ultraCheckEditor1.TabIndex = 0;
			this.ultraCheckEditor1.Text = "Contabilidad Gubernamental";
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.checkEdit11);
			this.ultraTabPageControl4.Controls.Add(this.checkEdit10);
			this.ultraTabPageControl4.Controls.Add(this.checkEdit8);
			this.ultraTabPageControl4.Controls.Add(this.checkEdit7);
			this.ultraTabPageControl4.Controls.Add(this.checkEdit5);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10769);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(612, 370);
			// 
			// checkEdit11
			// 
			this.checkEdit11.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FleteInventario"));
			this.checkEdit11.Location = new System.Drawing.Point(32, 164);
			this.checkEdit11.Name = "checkEdit11";
			// 
			// checkEdit11.Properties
			// 
			this.checkEdit11.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit11.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit11.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit11.Properties.Caption = "Incluir Flete en Inventario al ingresar Compra";
			this.checkEdit11.Size = new System.Drawing.Size(280, 21);
			this.checkEdit11.TabIndex = 79;
			// 
			// checkEdit10
			// 
			this.checkEdit10.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.IvaInventario"));
			this.checkEdit10.Location = new System.Drawing.Point(32, 129);
			this.checkEdit10.Name = "checkEdit10";
			// 
			// checkEdit10.Properties
			// 
			this.checkEdit10.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit10.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit10.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit10.Properties.Caption = "Incluir Iva en Inventario al ingresar compra";
			this.checkEdit10.Size = new System.Drawing.Size(280, 21);
			this.checkEdit10.TabIndex = 78;
			// 
			// checkEdit8
			// 
			this.checkEdit8.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.ImprimeIngreso"));
			this.checkEdit8.Location = new System.Drawing.Point(32, 95);
			this.checkEdit8.Name = "checkEdit8";
			// 
			// checkEdit8.Properties
			// 
			this.checkEdit8.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit8.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit8.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit8.Properties.Caption = "Banco en Impresión de Ingreso ";
			this.checkEdit8.Size = new System.Drawing.Size(280, 21);
			this.checkEdit8.TabIndex = 77;
			// 
			// checkEdit7
			// 
			this.checkEdit7.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaLinea"));
			this.checkEdit7.Location = new System.Drawing.Point(32, 60);
			this.checkEdit7.Name = "checkEdit7";
			// 
			// checkEdit7.Properties
			// 
			this.checkEdit7.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit7.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit7.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit7.Properties.Caption = "Estado de Cuenta por fecha de Línea de Asiento";
			this.checkEdit7.Size = new System.Drawing.Size(280, 21);
			this.checkEdit7.TabIndex = 76;
			// 
			// checkEdit5
			// 
			this.checkEdit5.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoC, "SeteoC.FechaDia"));
			this.checkEdit5.Location = new System.Drawing.Point(32, 26);
			this.checkEdit5.Name = "checkEdit5";
			// 
			// checkEdit5.Properties
			// 
			this.checkEdit5.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.checkEdit5.Properties.Appearance.Options.UseBackColor = true;
			this.checkEdit5.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit5.Properties.Caption = "Asientos de Facturas por fecha del Sistema y no por fecha de la factura";
			this.checkEdit5.Size = new System.Drawing.Size(400, 21);
			this.checkEdit5.TabIndex = 75;
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btAceptar.Location = new System.Drawing.Point(80, 422);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btCancelar.Location = new System.Drawing.Point(416, 422);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 22;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			this.btCancelar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCancelar_MouseUp);
			// 
			// ultraTabControl1
			// 
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(224)), ((System.Byte)(192)));
			this.ultraTabControl1.ActiveTabAppearance = appearance6;
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Location = new System.Drawing.Point(16, 9);
			this.ultraTabControl1.Name = "ultraTabControl1";
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.ultraTabControl1.SelectedTabAppearance = appearance7;
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(616, 396);
			this.ultraTabControl1.TabIndex = 23;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(224)), ((System.Byte)(192)));
			ultraTab1.ActiveAppearance = appearance8;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "General";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Numeración";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Gubernamental";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Predeterminados";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3,
																																															 ultraTab4});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(612, 370);
			// 
			// SeteoC
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(648, 471);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "SeteoC";
			this.Text = "Propiedades de Contabilidad";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.SeteoC_Closing);
			this.Load += new System.EventHandler(this.SeteoC_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.checkEdit9.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit6.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit4.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit3.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit2.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCapital.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentroCosto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCopiaBenef.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCopiaNota.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEgresos.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresos.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPasivo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaU.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtActivo.Properties)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoAsiento)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.checkEdit11.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit10.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit8.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit7.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit5.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsSeteoC.Update();
			Contabilidad.TipoCuenta(this.cdsCuenta, this.txtActivo.Text, 1);
			Contabilidad.TipoCuenta(this.cdsCuenta, this.txtPasivo.Text, 2);
			Contabilidad.TipoCuenta(this.cdsCuenta, this.txtCapital.Text, 3);
			Contabilidad.TipoCuenta(this.cdsCuenta, this.txtIngresos.Text, 4);
			Contabilidad.TipoCuenta(this.cdsCuenta, this.txtEgresos.Text, 5);
			cdsTipoAsiento.Update();
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		

		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsTipoAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipoAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void SeteoC_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Prop. Cont.'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsTipoAsiento, "Asiento", "idAsiento=0");
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			C1.Data.FilterCondition fcHija = new 
				C1.Data.FilterCondition(this.cdsTipoAsiento, "AsientoTipo", "idTipoAsiento > 0");
			fcTotal.Add(fcHija);

			C1.Data.FilterCondition fcHija1 = new 
				C1.Data.FilterCondition(this.cdsTipoAsiento, "AsientoDetalle", "idAsiento = 0");
			fcTotal.Add(fcHija1);
			cdsTipoAsiento.Fill(fcTotal, false);
		}

		private void SeteoC_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Prop. Fact.'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
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

		private void btListaCuenta_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea Actualizar Detalles de Asientos", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			Cursor = Cursors.WaitCursor;
			string stSelect = "Update AsientoDetalle Set AsientoDetalle.Codigo = Cuenta.Codigo "
				+ "From AsientoDetalle Inner Join Cuenta On AsientoDetalle.idCuenta = Cuenta.idCuenta ";
			Funcion.EjecutaSQL(cdsSeteoC, stSelect, true);
			stSelect = "Update AsientoDetalle Set AsientoDetalle.Descripcion = Cuenta.Descripcion "
				+ "From AsientoDetalle Inner Join Cuenta On AsientoDetalle.idCuenta = Cuenta.idCuenta ";
			Funcion.EjecutaSQL(cdsSeteoC, stSelect, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Datos Actualizados", "Información");
		}

		private void btCancelar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				Funcion.CreaFormulario(this);
		}

		private void btPlanCuenta_Click(object sender, System.EventArgs e)
		{
			using (PlanCuentas miPlan = new PlanCuentas())
			{
				miPlan.ShowDialog();
			}
			cdsCuenta.Clear();
			cdsCuenta.Fill();
		}


	}
}
