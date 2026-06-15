using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepFacturacion.
	/// </summary>
	public class RepFacturacion : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.RadioGroup rdgTipo;
		private C1.Data.C1DataSet cdsTipoFact;
		private C1.Data.C1DataView cdvTipoFact;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.LookUpEdit cmbTipoFact;
		private DevExpress.XtraEditors.ComboBoxEdit cmbFechas;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private DevExpress.XtraEditors.LookUpEdit cmbCliente;
		private DevExpress.XtraEditors.LookUpEdit cmbGrupo;
		private DevExpress.XtraEditors.CheckEdit chkGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbCentro;
		private DevExpress.XtraEditors.CheckEdit chkCentro;
		private DevExpress.XtraEditors.RadioGroup rdgFecha;
		private DevExpress.XtraEditors.LookUpEdit cmbComprobante;
		private DevExpress.XtraEditors.CheckEdit chkTipoTrans;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsGrupoCliente;
		private C1.Data.C1DataView cdvGrupoCliente;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbCiudad;
		private DevExpress.XtraEditors.CheckEdit chkCiudad;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.CheckEdit chkResumen;
		private DevExpress.XtraEditors.CheckEdit chkAnulado;
		private C1.Data.C1DataSet cdsCiudad;
		private C1.Data.C1DataView cdvCiudad;
		private C1.Data.C1DataSet cdsComprobante;
		private C1.Data.C1DataView cdvComprobante;
		private C1.Data.C1DataView cdvSector;
		private C1.Data.C1DataView cdvReporte;
		private DevExpress.XtraEditors.LookUpEdit cmbFormaPago;
		private DevExpress.XtraEditors.CheckEdit chkFormaPago;
		private C1.Data.C1DataSet cdsFormaPago;
		private C1.Data.C1DataView cdvFormaPago;
		private C1.Data.C1DataSet cdsSeteoRepF;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraEditors.SimpleButton btCierreCaja;
		private DevExpress.XtraEditors.CheckEdit chkVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsUsuario;
		private C1.Data.C1DataView cdvCodigoCliente;
		private C1.Data.C1DataSet cdsCompraTabla;
		private DevExpress.XtraEditors.CheckEdit chkOtro;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOtro;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsBodegas;
		private DevExpress.XtraEditors.CheckEdit chkRise;
		private DevExpress.XtraEditors.CheckEdit chkNeto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFacDesde;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFacHasta;
		private DevExpress.XtraEditors.CheckEdit chkRango;
		private DevExpress.XtraEditors.CheckEdit chkComprobante;
		private DevExpress.XtraEditors.CheckEdit chkCredTribut;
		private DevExpress.XtraEditors.LookUpEdit cmbVendedor;
		private DevExpress.XtraEditors.LookUpEdit cmbCredTribut;
		private DevExpress.XtraEditors.CheckEdit chkFiltroTipo;
		private DevExpress.XtraEditors.DateEdit dtFechaDesde;
		private DevExpress.XtraEditors.DateEdit dtFechaHasta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepFacturacion()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			this.rdgTipo = new DevExpress.XtraEditors.RadioGroup();
			this.cdsSeteoRepF = new C1.Data.C1DataSet();
			this.cmbTipoFact = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvTipoFact = new C1.Data.C1DataView();
			this.cdsTipoFact = new C1.Data.C1DataSet();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbFechas = new DevExpress.XtraEditors.ComboBoxEdit();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.cmbCliente = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvGrupoCliente = new C1.Data.C1DataView();
			this.cdsGrupoCliente = new C1.Data.C1DataSet();
			this.chkGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbCentro = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.chkCentro = new DevExpress.XtraEditors.CheckEdit();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.rdgFecha = new DevExpress.XtraEditors.RadioGroup();
			this.cmbComprobante = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvComprobante = new C1.Data.C1DataView();
			this.cdsComprobante = new C1.Data.C1DataSet();
			this.chkTipoTrans = new DevExpress.XtraEditors.CheckEdit();
			this.chkResumen = new DevExpress.XtraEditors.CheckEdit();
			this.chkAnulado = new DevExpress.XtraEditors.CheckEdit();
			this.cmbCiudad = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCiudad = new C1.Data.C1DataView();
			this.cdsCiudad = new C1.Data.C1DataSet();
			this.chkCiudad = new DevExpress.XtraEditors.CheckEdit();
			this.cdvSector = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.cmbFormaPago = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvFormaPago = new C1.Data.C1DataView();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.chkFormaPago = new DevExpress.XtraEditors.CheckEdit();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.udsUsuario = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btCierreCaja = new DevExpress.XtraEditors.SimpleButton();
			this.chkVendedor = new DevExpress.XtraEditors.CheckEdit();
			this.cdvCodigoCliente = new C1.Data.C1DataView();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.chkOtro = new DevExpress.XtraEditors.CheckEdit();
			this.txtOtro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.udsBodegas = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkRise = new DevExpress.XtraEditors.CheckEdit();
			this.chkNeto = new DevExpress.XtraEditors.CheckEdit();
			this.txtFacDesde = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFacHasta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkRango = new DevExpress.XtraEditors.CheckEdit();
			this.chkComprobante = new DevExpress.XtraEditors.CheckEdit();
			this.chkCredTribut = new DevExpress.XtraEditors.CheckEdit();
			this.cmbVendedor = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbCredTribut = new DevExpress.XtraEditors.LookUpEdit();
			this.chkFiltroTipo = new DevExpress.XtraEditors.CheckEdit();
			this.dtFechaDesde = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaHasta = new DevExpress.XtraEditors.DateEdit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRepF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoTrans.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumen.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnulado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCiudad.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFormaPago.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkVendedor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodigoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkOtro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRise.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNeto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRango.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCredTribut.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFiltroTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// rdgTipo
			// 
			this.rdgTipo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.NumReporte"));
			this.rdgTipo.EditValue = 0;
			this.rdgTipo.Location = new System.Drawing.Point(16, 17);
			this.rdgTipo.Name = "rdgTipo";
			// 
			// rdgTipo.Properties
			// 
			this.rdgTipo.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgTipo.Properties.Appearance.Options.UseBackColor = true;
			this.rdgTipo.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Tipo de Factura"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Saldos de Cuentas"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Abonos a Cuentas")});
			this.rdgTipo.Size = new System.Drawing.Size(160, 112);
			this.rdgTipo.TabIndex = 0;
			this.rdgTipo.Validated += new System.EventHandler(this.rdgTipo_SelectedIndexChanged);
			// 
			// cdsSeteoRepF
			// 
			this.cdsSeteoRepF.BindingContextCtrl = this;
			this.cdsSeteoRepF.DataLibrary = "LibFacturacion";
			this.cdsSeteoRepF.DataLibraryUrl = "";
			this.cdsSeteoRepF.DataSetDef = "dsSeteoF";
			this.cdsSeteoRepF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoRepF.Name = "cdsSeteoRepF";
			this.cdsSeteoRepF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoRepF.SchemaDef = null;
			this.cdsSeteoRepF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoRepF_BeforeFill);
			// 
			// cmbTipoFact
			// 
			this.cmbTipoFact.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Transaccion"));
			this.cmbTipoFact.Location = new System.Drawing.Point(384, 16);
			this.cmbTipoFact.Name = "cmbTipoFact";
			// 
			// cmbTipoFact.Properties
			// 
			this.cmbTipoFact.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipoFact.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbTipoFact.Properties.DataSource = this.cdvTipoFact;
			this.cmbTipoFact.Properties.DisplayMember = "Nombre";
			this.cmbTipoFact.Properties.NullText = "";
			this.cmbTipoFact.Properties.ShowFooter = false;
			this.cmbTipoFact.Properties.ShowHeader = false;
			this.cmbTipoFact.Properties.ValueMember = "idTipoFactura";
			this.cmbTipoFact.Size = new System.Drawing.Size(176, 20);
			this.cmbTipoFact.TabIndex = 1;
			// 
			// cdvTipoFact
			// 
			this.cdvTipoFact.BindingContextCtrl = this;
			this.cdvTipoFact.DataSet = this.cdsTipoFact;
			this.cdvTipoFact.RowFilter = "Grupo = 1 Or Grupo = 2 ";
			this.cdvTipoFact.TableName = "";
			this.cdvTipoFact.TableViewName = "CompraNumero";
			// 
			// cdsTipoFact
			// 
			this.cdsTipoFact.BindingContextCtrl = this;
			this.cdsTipoFact.DataLibrary = "LibFacturacion";
			this.cdsTipoFact.DataLibraryUrl = "";
			this.cdsTipoFact.DataSetDef = "dsCompraTabla";
			this.cdsTipoFact.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoFact.Name = "cdsTipoFact";
			this.cdsTipoFact.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoFact.SchemaDef = null;
			this.cdsTipoFact.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoFact_BeforeFill);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(384, 416);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 2;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(480, 416);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbFechas
			// 
			this.cmbFechas.EditValue = "Hasta Hoy";
			this.cmbFechas.Location = new System.Drawing.Point(16, 344);
			this.cmbFechas.Name = "cmbFechas";
			// 
			// cmbFechas.Properties
			// 
			this.cmbFechas.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbFechas.Properties.Items.AddRange(new object[] {
																															"Hasta Hoy",
																															"Solo Hoy",
																															"Este Mes",
																															"El mes pasado",
																															"Este Año",
																															"El Año pasado",
																															"Periodo Ingresado",
																															"Primer Trimestre",
																															"Segundo Trimestre",
																															"Tercer Trimestre",
																															"Cuarto Trimestre",
																															"Fechas Digitadas"});
			this.cmbFechas.Size = new System.Drawing.Size(184, 20);
			this.cmbFechas.TabIndex = 52;
			this.cmbFechas.SelectedIndexChanged += new System.EventHandler(this.cmbFechas_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 392);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 51;
			this.label2.Text = "Hasta:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 370);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 50;
			this.label1.Text = "Desde:";
			// 
			// chkCliente
			// 
			this.chkCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bCliente"));
			this.chkCliente.EditValue = true;
			this.chkCliente.Location = new System.Drawing.Point(232, 42);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.Caption = "Cliente";
			this.chkCliente.Size = new System.Drawing.Size(128, 18);
			this.chkCliente.TabIndex = 53;
			this.chkCliente.MouseDown += new System.Windows.Forms.MouseEventHandler(this.chkCliente_MouseDown);
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// cmbCliente
			// 
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Cliente"));
			this.cmbCliente.Location = new System.Drawing.Point(384, 41);
			this.cmbCliente.Name = "cmbCliente";
			// 
			// cmbCliente.Properties
			// 
			this.cmbCliente.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbCliente.Properties.DataSource = this.cdvCliente;
			this.cmbCliente.Properties.DisplayMember = "Nombre";
			this.cmbCliente.Properties.NullText = "";
			this.cmbCliente.Properties.PopupWidth = 300;
			this.cmbCliente.Properties.ShowFooter = false;
			this.cmbCliente.Properties.ShowHeader = false;
			this.cmbCliente.Properties.ValueMember = "idCliente";
			this.cmbCliente.Size = new System.Drawing.Size(176, 20);
			this.cmbCliente.TabIndex = 54;
			this.cmbCliente.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbCliente_MouseDown);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Grupo"));
			this.cmbGrupo.Location = new System.Drawing.Point(384, 66);
			this.cmbGrupo.Name = "cmbGrupo";
			// 
			// cmbGrupo.Properties
			// 
			this.cmbGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbGrupo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Grupo", "Grupo", 41, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbGrupo.Properties.DataSource = this.cdvGrupoCliente;
			this.cmbGrupo.Properties.DisplayMember = "Grupo";
			this.cmbGrupo.Properties.NullText = "";
			this.cmbGrupo.Properties.PopupWidth = 300;
			this.cmbGrupo.Properties.ShowFooter = false;
			this.cmbGrupo.Properties.ShowHeader = false;
			this.cmbGrupo.Properties.ValueMember = "idGrupoCliente";
			this.cmbGrupo.Size = new System.Drawing.Size(176, 20);
			this.cmbGrupo.TabIndex = 56;
			// 
			// cdvGrupoCliente
			// 
			this.cdvGrupoCliente.BindingContextCtrl = this;
			this.cdvGrupoCliente.DataSet = this.cdsGrupoCliente;
			this.cdvGrupoCliente.TableName = "";
			this.cdvGrupoCliente.TableViewName = "ClienteGrupo";
			// 
			// cdsGrupoCliente
			// 
			this.cdsGrupoCliente.BindingContextCtrl = this;
			this.cdsGrupoCliente.DataLibrary = "LibFacturacion";
			this.cdsGrupoCliente.DataLibraryUrl = "";
			this.cdsGrupoCliente.DataSetDef = "dsClienteTabla";
			this.cdsGrupoCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoCliente.Name = "cdsGrupoCliente";
			this.cdsGrupoCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoCliente.SchemaDef = null;
			this.cdsGrupoCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupoCliente_BeforeFill);
			// 
			// chkGrupo
			// 
			this.chkGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bGrupo"));
			this.chkGrupo.EditValue = true;
			this.chkGrupo.Location = new System.Drawing.Point(232, 67);
			this.chkGrupo.Name = "chkGrupo";
			// 
			// chkGrupo.Properties
			// 
			this.chkGrupo.Properties.Caption = "Grupo Cliente";
			this.chkGrupo.Size = new System.Drawing.Size(136, 18);
			this.chkGrupo.TabIndex = 55;
			this.chkGrupo.CheckedChanged += new System.EventHandler(this.chkGrupo_CheckedChanged);
			// 
			// cmbCentro
			// 
			this.cmbCentro.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Centro"));
			this.cmbCentro.Location = new System.Drawing.Point(384, 91);
			this.cmbCentro.Name = "cmbCentro";
			// 
			// cmbCentro.Properties
			// 
			this.cmbCentro.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCentro.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCentro.Properties.DataSource = this.cdvProyecto;
			this.cmbCentro.Properties.DisplayMember = "Nombre";
			this.cmbCentro.Properties.NullText = "";
			this.cmbCentro.Properties.PopupWidth = 300;
			this.cmbCentro.Properties.ShowFooter = false;
			this.cmbCentro.Properties.ShowHeader = false;
			this.cmbCentro.Properties.ValueMember = "idProyecto";
			this.cmbCentro.Size = new System.Drawing.Size(176, 20);
			this.cmbCentro.TabIndex = 60;
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
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// chkCentro
			// 
			this.chkCentro.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bCentro"));
			this.chkCentro.EditValue = true;
			this.chkCentro.Location = new System.Drawing.Point(232, 92);
			this.chkCentro.Name = "chkCentro";
			// 
			// chkCentro.Properties
			// 
			this.chkCentro.Properties.Caption = "BODEGA";
			this.chkCentro.Size = new System.Drawing.Size(120, 18);
			this.chkCentro.TabIndex = 59;
			this.chkCentro.MouseDown += new System.Windows.Forms.MouseEventHandler(this.chkCentro_MouseDown);
			this.chkCentro.CheckedChanged += new System.EventHandler(this.chkCentro_CheckedChanged);
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// rdgFecha
			// 
			this.rdgFecha.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.TipoFecha"));
			this.rdgFecha.EditValue = 0;
			this.rdgFecha.Location = new System.Drawing.Point(16, 258);
			this.rdgFecha.Name = "rdgFecha";
			// 
			// rdgFecha.Properties
			// 
			this.rdgFecha.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgFecha.Properties.Appearance.Options.UseBackColor = true;
			this.rdgFecha.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Emisión"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Vencimiento"),
																																																		 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Entrega")});
			this.rdgFecha.Size = new System.Drawing.Size(96, 78);
			this.rdgFecha.TabIndex = 63;
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Comprobante"));
			this.cmbComprobante.Location = new System.Drawing.Point(384, 117);
			this.cmbComprobante.Name = "cmbComprobante";
			// 
			// cmbComprobante.Properties
			// 
			this.cmbComprobante.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbComprobante.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Comprobante", "Comprobante", 78, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbComprobante.Properties.DataSource = this.cdvComprobante;
			this.cmbComprobante.Properties.DisplayMember = "Comprobante";
			this.cmbComprobante.Properties.NullText = "";
			this.cmbComprobante.Properties.PopupWidth = 300;
			this.cmbComprobante.Properties.ShowFooter = false;
			this.cmbComprobante.Properties.ShowHeader = false;
			this.cmbComprobante.Properties.ValueMember = "idComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(176, 20);
			this.cmbComprobante.TabIndex = 65;
			// 
			// cdvComprobante
			// 
			this.cdvComprobante.BindingContextCtrl = this;
			this.cdvComprobante.DataSet = this.cdsComprobante;
			this.cdvComprobante.TableName = "";
			this.cdvComprobante.TableViewName = "CompraComprobante";
			// 
			// cdsComprobante
			// 
			this.cdsComprobante.BindingContextCtrl = this;
			this.cdsComprobante.DataLibrary = "LibFacturacion";
			this.cdsComprobante.DataLibraryUrl = "";
			this.cdsComprobante.DataSetDef = "dsCompraTabla";
			this.cdsComprobante.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsComprobante.Name = "cdsComprobante";
			this.cdsComprobante.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsComprobante.SchemaDef = null;
			this.cdsComprobante.BeforeFill += new C1.Data.FillEventHandler(this.cdsComprobante_BeforeFill);
			// 
			// chkTipoTrans
			// 
			this.chkTipoTrans.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bTransaccion"));
			this.chkTipoTrans.Location = new System.Drawing.Point(232, 17);
			this.chkTipoTrans.Name = "chkTipoTrans";
			// 
			// chkTipoTrans.Properties
			// 
			this.chkTipoTrans.Properties.Caption = "Tipo Transacción";
			this.chkTipoTrans.Size = new System.Drawing.Size(144, 18);
			this.chkTipoTrans.TabIndex = 66;
			this.chkTipoTrans.CheckedChanged += new System.EventHandler(this.chkTipoTrans_CheckedChanged);
			// 
			// chkResumen
			// 
			this.chkResumen.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Resumen"));
			this.chkResumen.Location = new System.Drawing.Point(24, 152);
			this.chkResumen.Name = "chkResumen";
			// 
			// chkResumen.Properties
			// 
			this.chkResumen.Properties.Caption = "Resumen";
			this.chkResumen.Size = new System.Drawing.Size(75, 18);
			this.chkResumen.TabIndex = 67;
			// 
			// chkAnulado
			// 
			this.chkAnulado.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Anulados"));
			this.chkAnulado.Location = new System.Drawing.Point(24, 184);
			this.chkAnulado.Name = "chkAnulado";
			// 
			// chkAnulado.Properties
			// 
			this.chkAnulado.Properties.Caption = "Mostrar Anulados";
			this.chkAnulado.Size = new System.Drawing.Size(128, 18);
			this.chkAnulado.TabIndex = 68;
			// 
			// cmbCiudad
			// 
			this.cmbCiudad.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.Ciudad"));
			this.cmbCiudad.Location = new System.Drawing.Point(384, 142);
			this.cmbCiudad.Name = "cmbCiudad";
			// 
			// cmbCiudad.Properties
			// 
			this.cmbCiudad.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCiudad.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Ciudad", "Ciudad", 58, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCiudad.Properties.DataSource = this.cdvCiudad;
			this.cmbCiudad.Properties.DisplayMember = "Ciudad";
			this.cmbCiudad.Properties.NullText = "";
			this.cmbCiudad.Properties.PopupWidth = 300;
			this.cmbCiudad.Properties.ShowFooter = false;
			this.cmbCiudad.Properties.ShowHeader = false;
			this.cmbCiudad.Properties.ValueMember = "idCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(176, 20);
			this.cmbCiudad.TabIndex = 70;
			// 
			// cdvCiudad
			// 
			this.cdvCiudad.BindingContextCtrl = this;
			this.cdvCiudad.DataSet = this.cdsCiudad;
			this.cdvCiudad.Sort = "Ciudad";
			this.cdvCiudad.TableName = "";
			this.cdvCiudad.TableViewName = "ClienteCiudad";
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
			// chkCiudad
			// 
			this.chkCiudad.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bCiudad"));
			this.chkCiudad.EditValue = true;
			this.chkCiudad.Location = new System.Drawing.Point(232, 143);
			this.chkCiudad.Name = "chkCiudad";
			// 
			// chkCiudad.Properties
			// 
			this.chkCiudad.Properties.Caption = "Ciudad";
			this.chkCiudad.Size = new System.Drawing.Size(120, 18);
			this.chkCiudad.TabIndex = 69;
			this.chkCiudad.MouseDown += new System.Windows.Forms.MouseEventHandler(this.chkCiudad_MouseDown);
			this.chkCiudad.CheckedChanged += new System.EventHandler(this.chkCiudad_CheckedChanged);
			// 
			// cdvSector
			// 
			this.cdvSector.BindingContextCtrl = this;
			this.cdvSector.DataSet = this.cdsCiudad;
			this.cdvSector.TableName = "";
			this.cdvSector.TableViewName = "ClienteSector";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
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
			this.cdsReporte.BeforeFill += new C1.Data.FillEventHandler(this.cdsReporte_BeforeFill);
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.FormaPago"));
			this.cmbFormaPago.Location = new System.Drawing.Point(384, 168);
			this.cmbFormaPago.Name = "cmbFormaPago";
			// 
			// cmbFormaPago.Properties
			// 
			this.cmbFormaPago.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbFormaPago.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("FormaPago", "FormaPago", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbFormaPago.Properties.DataSource = this.cdvFormaPago;
			this.cmbFormaPago.Properties.DisplayMember = "FormaPago";
			this.cmbFormaPago.Properties.NullText = "";
			this.cmbFormaPago.Properties.PopupWidth = 300;
			this.cmbFormaPago.Properties.ShowFooter = false;
			this.cmbFormaPago.Properties.ShowHeader = false;
			this.cmbFormaPago.Properties.ValueMember = "idFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 20);
			this.cmbFormaPago.TabIndex = 77;
			// 
			// cdvFormaPago
			// 
			this.cdvFormaPago.BindingContextCtrl = this;
			this.cdvFormaPago.DataSet = this.cdsFormaPago;
			this.cdvFormaPago.TableName = "";
			this.cdvFormaPago.TableViewName = "CompraFormaPago";
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
			// 
			// chkFormaPago
			// 
			this.chkFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bFormaPago"));
			this.chkFormaPago.EditValue = true;
			this.chkFormaPago.Location = new System.Drawing.Point(232, 169);
			this.chkFormaPago.Name = "chkFormaPago";
			// 
			// chkFormaPago.Properties
			// 
			this.chkFormaPago.Properties.Caption = "Forma de Pago";
			this.chkFormaPago.Size = new System.Drawing.Size(136, 18);
			this.chkFormaPago.TabIndex = 76;
			this.chkFormaPago.MouseDown += new System.Windows.Forms.MouseEventHandler(this.chkFormaPago_MouseDown);
			this.chkFormaPago.CheckedChanged += new System.EventHandler(this.chkFormaPago_CheckedChanged);
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.RowFilter = "Vendedor = 1";
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// udsUsuario
			// 
			this.udsUsuario.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1});
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoRepF, "SeteoRepFact.TipoReporte"));
			this.cmbReporte.DataSource = this.cdvReporte;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 8;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
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
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(16, 224);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(184, 21);
			this.cmbReporte.TabIndex = 92;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("ro-RO");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// btCierreCaja
			// 
			this.btCierreCaja.Location = new System.Drawing.Point(496, 304);
			this.btCierreCaja.Name = "btCierreCaja";
			this.btCierreCaja.Size = new System.Drawing.Size(75, 25);
			this.btCierreCaja.TabIndex = 93;
			this.btCierreCaja.Text = "Cierre Ca&ja";
			this.btCierreCaja.Visible = false;
			this.btCierreCaja.Click += new System.EventHandler(this.btCierreCaja_Click);
			this.btCierreCaja.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCierreCaja_MouseUp);
			// 
			// chkVendedor
			// 
			this.chkVendedor.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bVendedor"));
			this.chkVendedor.EditValue = true;
			this.chkVendedor.Location = new System.Drawing.Point(232, 193);
			this.chkVendedor.Name = "chkVendedor";
			// 
			// chkVendedor.Properties
			// 
			this.chkVendedor.Properties.Caption = "Vendedor";
			this.chkVendedor.Size = new System.Drawing.Size(112, 18);
			this.chkVendedor.TabIndex = 94;
			this.chkVendedor.MouseDown += new System.Windows.Forms.MouseEventHandler(this.chkVendedor_MouseDown);
			this.chkVendedor.CheckedChanged += new System.EventHandler(this.chkVendedor_CheckedChanged);
			// 
			// cdvCodigoCliente
			// 
			this.cdvCodigoCliente.BindingContextCtrl = this;
			this.cdvCodigoCliente.DataSet = this.cdsCliente;
			this.cdvCodigoCliente.Sort = "Codigo";
			this.cdvCodigoCliente.TableName = "";
			this.cdvCodigoCliente.TableViewName = "Cliente";
			// 
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraTabla_BeforeFill);
			// 
			// chkOtro
			// 
			this.chkOtro.EditValue = true;
			this.chkOtro.Location = new System.Drawing.Point(384, 264);
			this.chkOtro.Name = "chkOtro";
			// 
			// chkOtro.Properties
			// 
			this.chkOtro.Properties.Caption = "Otro";
			this.chkOtro.Size = new System.Drawing.Size(56, 18);
			this.chkOtro.TabIndex = 102;
			this.chkOtro.CheckedChanged += new System.EventHandler(this.chkOtro_CheckedChanged);
			// 
			// txtOtro
			// 
			this.txtOtro.Enabled = false;
			this.txtOtro.Location = new System.Drawing.Point(384, 288);
			this.txtOtro.Name = "txtOtro";
			this.txtOtro.Size = new System.Drawing.Size(72, 22);
			this.txtOtro.TabIndex = 103;
			// 
			// udsBodegas
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsBodegas.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn2,
																													 ultraDataColumn3});
			// 
			// chkRise
			// 
			this.chkRise.Location = new System.Drawing.Point(240, 264);
			this.chkRise.Name = "chkRise";
			// 
			// chkRise.Properties
			// 
			this.chkRise.Properties.Caption = "Rise o Inst. Publica";
			this.chkRise.Size = new System.Drawing.Size(120, 18);
			this.chkRise.TabIndex = 128;
			// 
			// chkNeto
			// 
			this.chkNeto.Location = new System.Drawing.Point(120, 152);
			this.chkNeto.Name = "chkNeto";
			// 
			// chkNeto.Properties
			// 
			this.chkNeto.Properties.Caption = "Neto";
			this.chkNeto.Size = new System.Drawing.Size(75, 18);
			this.chkNeto.TabIndex = 129;
			// 
			// txtFacDesde
			// 
			this.txtFacDesde.Enabled = false;
			this.txtFacDesde.Location = new System.Drawing.Point(128, 284);
			this.txtFacDesde.Name = "txtFacDesde";
			this.txtFacDesde.Size = new System.Drawing.Size(72, 22);
			this.txtFacDesde.TabIndex = 130;
			this.txtFacDesde.Text = "1";
			this.txtFacDesde.Visible = false;
			// 
			// txtFacHasta
			// 
			this.txtFacHasta.Enabled = false;
			this.txtFacHasta.Location = new System.Drawing.Point(128, 310);
			this.txtFacHasta.Name = "txtFacHasta";
			this.txtFacHasta.Size = new System.Drawing.Size(72, 22);
			this.txtFacHasta.TabIndex = 131;
			this.txtFacHasta.Text = "999";
			this.txtFacHasta.Visible = false;
			// 
			// chkRango
			// 
			this.chkRango.EditValue = true;
			this.chkRango.Location = new System.Drawing.Point(128, 258);
			this.chkRango.Name = "chkRango";
			// 
			// chkRango.Properties
			// 
			this.chkRango.Properties.Caption = "Rango";
			this.chkRango.Size = new System.Drawing.Size(64, 18);
			this.chkRango.TabIndex = 132;
			this.chkRango.Visible = false;
			this.chkRango.CheckedChanged += new System.EventHandler(this.chkRango_CheckedChanged);
			// 
			// chkComprobante
			// 
			this.chkComprobante.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.bComprobante"));
			this.chkComprobante.EditValue = true;
			this.chkComprobante.Location = new System.Drawing.Point(232, 118);
			this.chkComprobante.Name = "chkComprobante";
			// 
			// chkComprobante.Properties
			// 
			this.chkComprobante.Properties.Caption = "Comprobante";
			this.chkComprobante.Size = new System.Drawing.Size(136, 18);
			this.chkComprobante.TabIndex = 135;
			this.chkComprobante.MouseDown += new System.Windows.Forms.MouseEventHandler(this.chkComprobante_MouseDown);
			this.chkComprobante.CheckedChanged += new System.EventHandler(this.chkComprobante_CheckedChanged);
			// 
			// chkCredTribut
			// 
			this.chkCredTribut.EditValue = true;
			this.chkCredTribut.Location = new System.Drawing.Point(232, 219);
			this.chkCredTribut.Name = "chkCredTribut";
			// 
			// chkCredTribut.Properties
			// 
			this.chkCredTribut.Properties.Caption = "Todo Crédito Tributario";
			this.chkCredTribut.Size = new System.Drawing.Size(152, 18);
			this.chkCredTribut.TabIndex = 136;
			this.chkCredTribut.MouseDown += new System.Windows.Forms.MouseEventHandler(this.chkCredTribut_MouseDown);
			this.chkCredTribut.CheckedChanged += new System.EventHandler(this.chkCredTribut_CheckedChanged);
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.Location = new System.Drawing.Point(384, 192);
			this.cmbVendedor.Name = "cmbVendedor";
			// 
			// cmbVendedor.Properties
			// 
			this.cmbVendedor.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbVendedor.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbVendedor.Properties.DisplayMember = "Nombre";
			this.cmbVendedor.Properties.NullText = "";
			this.cmbVendedor.Properties.PopupWidth = 300;
			this.cmbVendedor.Properties.ShowFooter = false;
			this.cmbVendedor.Properties.ShowHeader = false;
			this.cmbVendedor.Properties.ValueMember = "idPersonal";
			this.cmbVendedor.Size = new System.Drawing.Size(176, 20);
			this.cmbVendedor.TabIndex = 138;
			// 
			// cmbCredTribut
			// 
			this.cmbCredTribut.Location = new System.Drawing.Point(384, 218);
			this.cmbCredTribut.Name = "cmbCredTribut";
			// 
			// cmbCredTribut.Properties
			// 
			this.cmbCredTribut.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCredTribut.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CredTributario", "CredTributario", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbCredTribut.Properties.DisplayMember = "CredTributario";
			this.cmbCredTribut.Properties.NullText = "";
			this.cmbCredTribut.Properties.PopupWidth = 300;
			this.cmbCredTribut.Properties.ShowFooter = false;
			this.cmbCredTribut.Properties.ShowHeader = false;
			this.cmbCredTribut.Properties.ValueMember = "idCredTributario";
			this.cmbCredTribut.Size = new System.Drawing.Size(176, 20);
			this.cmbCredTribut.TabIndex = 141;
			// 
			// chkFiltroTipo
			// 
			this.chkFiltroTipo.Location = new System.Drawing.Point(472, 264);
			this.chkFiltroTipo.Name = "chkFiltroTipo";
			// 
			// chkFiltroTipo.Properties
			// 
			this.chkFiltroTipo.Properties.Caption = "SoloFiltro";
			this.chkFiltroTipo.Size = new System.Drawing.Size(75, 18);
			this.chkFiltroTipo.TabIndex = 142;
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.FechaIni"));
			this.dtFechaDesde.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaDesde.Location = new System.Drawing.Point(104, 368);
			this.dtFechaDesde.Name = "dtFechaDesde";
			// 
			// dtFechaDesde.Properties
			// 
			this.dtFechaDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaDesde.Size = new System.Drawing.Size(96, 20);
			this.dtFechaDesde.TabIndex = 143;
			this.dtFechaDesde.EditValueChanged += new System.EventHandler(this.dtFechaDesde_EditValueChanged);
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoRepF, "SeteoRepFact.FechaFin"));
			this.dtFechaHasta.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaHasta.Location = new System.Drawing.Point(104, 392);
			this.dtFechaHasta.Name = "dtFechaHasta";
			// 
			// dtFechaHasta.Properties
			// 
			this.dtFechaHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFechaHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFechaHasta.Size = new System.Drawing.Size(96, 20);
			this.dtFechaHasta.TabIndex = 144;
			// 
			// RepFacturacion
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(584, 462);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.chkFiltroTipo);
			this.Controls.Add(this.cmbCredTribut);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.chkCredTribut);
			this.Controls.Add(this.chkComprobante);
			this.Controls.Add(this.chkRango);
			this.Controls.Add(this.txtFacHasta);
			this.Controls.Add(this.txtFacDesde);
			this.Controls.Add(this.chkNeto);
			this.Controls.Add(this.chkRise);
			this.Controls.Add(this.txtOtro);
			this.Controls.Add(this.chkOtro);
			this.Controls.Add(this.chkVendedor);
			this.Controls.Add(this.btCierreCaja);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.chkFormaPago);
			this.Controls.Add(this.cmbCiudad);
			this.Controls.Add(this.chkCiudad);
			this.Controls.Add(this.chkAnulado);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.chkTipoTrans);
			this.Controls.Add(this.cmbComprobante);
			this.Controls.Add(this.rdgFecha);
			this.Controls.Add(this.cmbCentro);
			this.Controls.Add(this.chkCentro);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.chkGrupo);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.cmbFechas);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbTipoFact);
			this.Controls.Add(this.rdgTipo);
			this.MaximizeBox = false;
			this.Name = "RepFacturacion";
			this.Text = "Reportes de Facturación";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepFacturacion_Closing);
			this.Load += new System.EventHandler(this.RepFacturacion_Load);
			this.Activated += new System.EventHandler(this.RepFacturacion_Activated);
			((System.ComponentModel.ISupportInitialize)(this.rdgTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRepF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoTrans.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumen.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnulado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCiudad.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFormaPago.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkVendedor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodigoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkOtro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRise.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNeto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkRango.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCredTribut.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkFiltroTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private bool Verifica()
		{
			bool bOk = true;
			this.errorProvider.SetError(this.cmbTipoFact, "");
			this.errorProvider.SetError(this.cmbCentro, "");
			this.errorProvider.SetError(this.cmbCiudad, "");
			this.errorProvider.SetError(this.cmbCliente, "");
			this.errorProvider.SetError(this.cmbComprobante, "");
			this.errorProvider.SetError(this.cmbGrupo, "");
			this.errorProvider.SetError(this.cmbFormaPago, "");
			this.errorProvider.SetError(this.cmbReporte, "");

			this.cmbReporte.Refresh();
			if (this.cmbReporte.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.cmbReporte, "Seleccione el reporte deseado");
				bOk = false;
			}

			if ((int) this.rdgTipo.EditValue > 0)
			{
				
				if (! ((int) this.cmbTipoFact.EditValue == 1 || (int) this.cmbTipoFact.EditValue == 4    ) )
				{
					this.errorProvider.SetError(this.cmbTipoFact, "Debe escoger compra o venta");
					bOk = false;
				}
				if (chkFiltroTipo.Checked == true)
				{
					bOk = true;
					this.errorProvider.SetError(this.cmbTipoFact,"");
				}
			}
			

			if (chkTipoTrans.Checked == false || chkTipoTrans.CheckState == CheckState.Indeterminate)
				if (cmbTipoFact.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbTipoFact, "Ingrese tipo de Factura");
					bOk = false;
				}
					
			if (chkCliente.Checked == false || chkCliente.CheckState == CheckState.Indeterminate)
				if (cmbCliente.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCliente, "Ingrese Cliente");
					bOk = false;
				}
			if (chkGrupo.Checked == false || chkGrupo.CheckState == CheckState.Indeterminate)
				if (cmbGrupo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbGrupo, "Ingrese Grupo de Cliente");
					bOk = false;
				}
			if (chkVendedor.Checked == false || chkVendedor.CheckState == CheckState.Indeterminate)
				if (cmbVendedor.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbVendedor, "Ingrese Vendedor");
					bOk = false;
				}
			if (chkCentro.Checked == false || chkCentro.CheckState == CheckState.Indeterminate)
				if (cmbCentro.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCentro, "Ingrese Centro de Costo");
					bOk = false;
				}
			if (chkComprobante.Checked == false || chkComprobante.CheckState == CheckState.Indeterminate)
				if (cmbComprobante.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbComprobante, "Ingrese Comprobante");
					bOk = false;
				}
			if (chkCiudad.Checked == false || chkCiudad.CheckState == CheckState.Indeterminate)
				if (cmbCiudad.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCiudad, "Ingrese Ciudad");
					bOk = false;
				}
			if (chkFormaPago.Checked == false || chkFormaPago.CheckState == CheckState.Indeterminate)
				if (cmbFormaPago.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbFormaPago, "Ingrese Forma de Pago");
					bOk = false;
				}
			if (chkCredTribut.Checked == false || chkCredTribut.CheckState == CheckState.Indeterminate)
				if (cmbCredTribut.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCredTribut, "Ingrese Receptor");
					bOk = false;
				}
			return bOk;
		}

		private string[] Filtro()
		{
			DateTime dtFechaIni = (DateTime) this.dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) this.dtFechaHasta.EditValue;

			string stFiltro = "";
			string stFiltroSerie = "";
			string stTitulo = "";
			string stFecha = "Fecha";
			if ((int) this.rdgFecha.EditValue == 1) stFecha = "FechaVencimiento";
			if ((int) this.rdgFecha.EditValue == 2) stFecha = "FechaEntrega";

			stFiltro = "{Compra." + stFecha + "} >= #"
				+ dtFechaIni.ToString("MM/dd/yyyy", us) 
				+ "# And {Compra." + stFecha + "} < #" 
				+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
			stTitulo = " Desde " + dtFechaIni.ToString("dd/MMM/yyyy") +
				" Hasta " + dtFechaFin.ToString("dd/MMM/yyyy") + " ";
			if ((int) rdgTipo.EditValue == 1)
			{
				string stFechaMax = dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
				stFiltro = "{Compra." + stFecha + "} < #" + stFechaMax;
				//+ " And ( ({Pago.Fecha} < #" + stFechaMax+ ") Or ({Pago.Fecha} Is NULL )) ";
				stTitulo = " Hasta " + dtFechaFin.ToString("dd/MMM/yyyy") + " ";
			}
			if ((int) rdgTipo.EditValue == 2) // Abonos a CxC o P
			{
				string stFechaMax = dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
				if ((int) rdgFecha.EditValue == 0)
				{
					stFiltro = "{Pago.Fecha} < #" + stFechaMax + 
						" And {Pago.Fecha} >= #" + dtFechaIni.ToString("MM/dd/yyyy", us) + "#";
				}
				else
				{
					stFiltro = "{Pago.Vencimiento} < #" + stFechaMax + 
						" And {Pago.Vencimiento} >= #" + dtFechaIni.ToString("MM/dd/yyyy", us) + "#";
				}
			}

			if (chkTipoTrans.Checked == false && (int) rdgTipo.EditValue != 1)
			{
				int iTipoF = int.Parse(cmbTipoFact.EditValue.ToString());
				if (chkNeto.Checked && (iTipoF == 1 || iTipoF == 4)) 
				{
					if (iTipoF == 1)
						stFiltro += " And {Compra.idTipoFactura} In [1, 5]";
					else
						stFiltro += " And {Compra.idTipoFactura} In [4, 6]";

//					stTitulo += "Tipo " + cmbTipoFact.Text.Trim() + " y Dev. ";
				}
				else
				{
					stFiltro += " And {Compra.idTipoFactura} = " + cmbTipoFact.EditValue.ToString();
//					stTitulo += "Tipo " + cmbTipoFact.Text.Trim() + " ";
				}
			}

			if (chkTipoTrans.Checked == false && (int) rdgTipo.EditValue == 1)
			{
				int iTipoF = int.Parse(this.cmbTipoFact.EditValue.ToString());
				if(chkFiltroTipo.Checked == true)
				{
						stFiltro += " And {Compra.idTipoFactura} = "+iTipoF;
				}
				else
				{
					if (iTipoF == 1)
					{
						stFiltro += " And {Compra.idTipoFactura} In [1, 5, 27, 11, 99] ";
					}
					else
					{
						stFiltro += " And {Compra.idTipoFactura} In [4, 6, 28, 26, 14] ";
					}
				}
			}

			#region Visto en Clientes
			if (chkCliente.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Compra.idCliente} <> " + cmbCliente.EditValue.ToString();
				stFiltroSerie += " And Compra.idCliente <> " + cmbCliente.EditValue.ToString();
				stTitulo += "Cliente No es " + cmbCliente.Text.Trim() + " ";
			}
			if (chkCliente.Checked == false)
			{
				stFiltro += " And {Compra.idCliente} = " + cmbCliente.EditValue.ToString();
				stFiltroSerie += " And Compra.idCliente = " + cmbCliente.EditValue.ToString();
				stTitulo += "Cliente " + cmbCliente.Text.Trim() + " ";
			}
			#endregion Visto en Cliente

			#region Visto en Grupo
			if (chkGrupo.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Cliente.idGrupoCliente} <> " + cmbGrupo.EditValue.ToString();
				stFiltroSerie += " And Cl.idGrupoCliente <> " + cmbGrupo.EditValue.ToString();
				stTitulo += "Grupo No es " + cmbGrupo.Text.Trim() + " ";
			}
			if (chkGrupo.Checked == false)
			{
				stFiltro += " And {Cliente.idGrupoCliente} = " + cmbGrupo.EditValue.ToString();
				stFiltroSerie += " And Cl.idGrupoCliente = " + cmbGrupo.EditValue.ToString();
				stTitulo += "Grupo " + cmbGrupo.Text.Trim() + " ";
			}
			#endregion Visto en Grupo

			if (chkRise.Checked == true)
			{
				stFiltro += " And {Cliente.InstPublica} = true ";
				stFiltroSerie += " And Cl.InstPublica = 1";
				stTitulo += "Rise ";
			}
			
			#region Vendedor
			if (chkVendedor.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'FCVB'")) 	
				{
					stFiltro += " And {DetCompra.idVendedor} <> " + cmbVendedor.EditValue.ToString();
					stFiltroSerie += " And DetCompra.idVendedor <> " + cmbVendedor.EditValue.ToString();
					stTitulo += "Vendedor No es " + cmbVendedor.Text.Trim() + " ";
				}
				else
				{
					stFiltro += " And {Compra.idVendedor} <> " + cmbVendedor.EditValue.ToString();
					stFiltroSerie += " And Compra.idVendedor <> " + cmbVendedor.EditValue.ToString();
					stTitulo += "Vendedor No es " + cmbVendedor.Text.Trim() + " ";
				}
			}

			if (chkVendedor.Checked == false)
			{
				if (Funcion.bEjecutaSQL(cdsCliente, "Exec SeteoGExiste 'FCVB'")) 	
				{
					stFiltro += " And {DetCompra.idVendedor} = " + cmbVendedor.EditValue.ToString();
					stFiltroSerie += " And DetCompra.idVendedor = " + cmbVendedor.EditValue.ToString();
					stTitulo += "Vendedor " + cmbVendedor.Text.Trim() + " ";
				}
				else
				{
					stFiltro += " And {Compra.idVendedor} = " + cmbVendedor.EditValue.ToString();
					stFiltroSerie += " And Compra.idVendedor = " + cmbVendedor.EditValue.ToString();
					stTitulo += "Vendedor " + cmbVendedor.Text.Trim() + " ";
				}
			}
			#endregion Vendedor

			#region Visto en Credito Tribitario
			if (chkCredTribut.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Compra.idCredTributario} <> " + cmbCredTribut.EditValue.ToString();
				stFiltroSerie += " And Compra.idCredTributario <> " + cmbCredTribut.EditValue.ToString();
				stTitulo += "Cred. Trib. No es " + cmbCredTribut.Text.Trim() + " ";
			}
			if (chkCredTribut.Checked == false)
			{
				stFiltro += " And {Compra.idCredTributario} = " + cmbCredTribut.EditValue.ToString();
				stFiltroSerie += " And Compra.idCredTributario = " + cmbCredTribut.EditValue.ToString();
				stTitulo += "Cred. Trib. " + cmbCredTribut.Text.Trim() + " ";
			}
			#endregion Visto en Credito Tribitario

			#region Visto en Centro de Costo
			if (chkCentro.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				if (drSeteoF["VerProyecto"] == DBNull.Value || (bool) drSeteoF["VerProyecto"] == true)
					stFiltro += " And {DetCompra.idProyecto} <> " + cmbCentro.EditValue.ToString();
				else
					stFiltro += " And {Compra.idProyecto} <> "	+ cmbCentro.EditValue.ToString();

				stFiltroSerie += " And Compra.idProyecto <> "	+ cmbCentro.EditValue.ToString();
				stTitulo += "Centro de Costo No es " + cmbCentro.Text.Trim() + " ";
			}
			if (chkCentro.Checked == false)
			{
				if (drSeteoF["VerProyecto"] == DBNull.Value || (bool) drSeteoF["VerProyecto"] == true)
					stFiltro += " And {DetCompra.idProyecto} = " + cmbCentro.EditValue.ToString();
				else
					stFiltro += " And {Compra.idProyecto} = "	+ cmbCentro.EditValue.ToString();

				stFiltroSerie += " And Compra.idProyecto = "	+ cmbCentro.EditValue.ToString();
				stTitulo += "Centro de Costo " + cmbCentro.Text.Trim() + " ";
			}
			#endregion Visto en Centro de Costo

			#region Visto en Comprobante
			if (chkComprobante.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Compra.idComprobante} <> " + cmbComprobante.EditValue.ToString();
				stFiltroSerie += " And Compra.idComprobante <> " + cmbComprobante.EditValue.ToString();
				stTitulo += "Comprobante No es " + cmbComprobante.Text.Trim() + " ";
			}
			if (chkComprobante.Checked == false)
			{
				stFiltro += " And {Compra.idComprobante} = " + cmbComprobante.EditValue.ToString();
				stFiltroSerie += " And Compra.idComprobante = " + cmbComprobante.EditValue.ToString();
				stTitulo += "Comprobante " + cmbComprobante.Text.Trim() + " ";
			}
			#endregion Visto en Comprobante

			#region Visto en Ciudad
			if (chkCiudad.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Cliente.idCiudad} <> " + cmbCiudad.EditValue.ToString();
				stFiltroSerie += " And Cl.idCiudad <> " + cmbCiudad.EditValue.ToString();
				stTitulo += "Ciudad No es " + cmbCiudad.Text.Trim() + " ";
			}
			if (chkCiudad.Checked == false)
			{
				stFiltro += " And {Cliente.idCiudad} = " + cmbCiudad.EditValue.ToString();
				stFiltroSerie += " And Cl.idCiudad = " + cmbCiudad.EditValue.ToString();
				stTitulo += "Ciudad " + cmbCiudad.Text.Trim() + " ";
			}
			#endregion Visto en Comprobante

			#region Forma de Pago
			if ((int) rdgTipo.EditValue == 2) // Abonos
			{
				if (this.chkFormaPago.Checked == false)
				{
					stFiltro += " And {Pago.idFormaPago} = " + cmbFormaPago.EditValue.ToString();
					stFiltroSerie += " And Pago.idFormaPago = " + cmbFormaPago.EditValue.ToString();
					stTitulo += "Forma de Pago " + cmbFormaPago.Text.Trim() + " ";
				}
			}
			else
			{
				if (chkFormaPago.Checked == false)
				{
					stFiltro += " And {Compra.idFormaPago} = " + cmbFormaPago.EditValue.ToString();
					stFiltroSerie += " And Compra.idFormaPago = " + cmbFormaPago.EditValue.ToString();
					stTitulo += "Forma de Pago " + cmbFormaPago.Text.Trim() + " ";
				}
			}
			#endregion Forma de Pago

			#region Anulado
			if (!this.chkAnulado.Checked)
			{
				stFiltro += " And {Compra.Borrar} = false" ;
			}
			#endregion Anulado

			#region Estado de cuenta
			if ((int) rdgTipo.EditValue == 1)
			{
				stFiltro += " And {Compra.ContadoCredito} = 2" ;
			}
			#endregion Estado de cuenta

			#region Visto en Otro
			if (chkOtro.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Compra.Otro} Not Like '" + txtOtro.Text + "'" ;
			}
			if (!chkOtro.Checked)
			{
				stFiltro += " And {Compra.Otro} Like '" + txtOtro.Text + "'" ;
			}
			#endregion Visto en Otro

			string[] stFiltro1 = new string[3];
			stFiltro1[0] = stFiltro;
			stFiltro1[1] = stTitulo;
			stFiltro1[2] = stFiltroSerie;
			return stFiltro1;
		}

		private string GeneraPeriodo()
		{
			DateTime dtFechaIni = (DateTime) this.dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) this.dtFechaHasta.EditValue;
			string stPeriodo = "Desde " + dtFechaIni.ToString("dd/MMM/yyyy")
			+ " Hasta " + dtFechaFin.ToString("dd/MMM/yyyy").Trim();

			string stFiltro = "";
			if (this.chkCliente.Checked == false)
			{
				stFiltro += " Cliente: " + this.cmbCliente.Text.Trim();
			}
			if (this.chkGrupo.Checked == false)
			{
				stFiltro += " Grupo "  + this.cmbGrupo.Text.Trim();
			}
			if (this.chkVendedor.Checked == false)
			{
				stFiltro += " Vendedor " + this.cmbVendedor.Text.Trim();
			}
			if (this.chkCentro.Checked == false)
			{
				stFiltro += " Centro: "+ this.cmbCentro.Text.Trim();
			}
			if (this.chkComprobante.Checked == false)
			{
				stFiltro += " Comprobante: " + this.cmbComprobante.Text.Trim();
			}
			if (this.chkCiudad.Checked == false)
			{
				stFiltro += " Ciudad: " + this.cmbCiudad.Text.Trim();
			}
			if (this.chkFormaPago.Checked == false)
			{
				stFiltro += " Forma de Pago: " + this.cmbFormaPago.Text.Trim();
			}
			if (this.chkAnulado.Checked)
			{
				stFiltro += " ANULADOS" ;
			}

			return stPeriodo + stFiltro;
		}

		private string GeneraTitulo()
		{
			string strTitulo = "";
			if ((int) this.rdgTipo.EditValue == 0)
			{
				strTitulo = this.cmbTipoFact.Text.ToString();
			}
			else if ((int) this.rdgTipo.EditValue == 1)
			{
				if ((int) this.cmbTipoFact.EditValue == 1)
					strTitulo = "Cuentas por Cobrar";
				else
					strTitulo = "Cuentas por Pagar";
			}
			else if ((int) this.rdgTipo.EditValue == 2)
			{
				if ((int) this.cmbTipoFact.EditValue == 1)
					strTitulo = "Abonos Entregados";
				else
					strTitulo = "Abonos Recibidos";
			}
			return strTitulo;
		}

		private bool Seguridad()
		{
			bool bOk = true;
			#region Filtro de Facturas
			if ((int) rdgTipo.EditValue == 0)
			{
				switch((int) cmbTipoFact.EditValue) // Factura de venta
				{
					case 1:
						if (!Funcion.Permiso("107", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Facturas de Venta");
							bOk = false;
						}
						break;
					case 2:
						if (!Funcion.Permiso("108", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Ordenes de Compra");
							bOk = false;
						}
						break;
					case 3:
						if (!Funcion.Permiso("109", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso de Cotizaciones");
							bOk = false;
						}
						break;
					case 4:
						if (!Funcion.Permiso("110", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Facturas de Compra");
							bOk = false;
						}
						break;
					case 5:
						if (!Funcion.Permiso("111", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Devolucion de Venta");
							bOk = false;
						}
						break;
					case 6:
						if (!Funcion.Permiso("112", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Devolucion de Compra");
							bOk = false;
						}
						break;
					case 8:
						if (!Funcion.Permiso("120", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Egresos de Bodega");
							bOk = false;
						}
						break;
					case 9:
						if (!Funcion.Permiso("121", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Ingresos de Bodega");
							bOk = false;
						}
						break;
				} // Fin de Switch
			} // Fin de If
			#endregion Fin de Filtro de Facturas

			#region Filtro de Cuentas por Cobrar y Pagar
			if ((int) rdgTipo.EditValue == 1 )
			{
				#region No se permite mirar todo cuando hay restriccciones en cxc o cxp
				if (chkTipoTrans.Checked)
				{
					if (!Funcion.Permiso("113", cdsSeguridad))
					{
						errorProvider.SetError(cmbTipoFact, "No tiene acceso a Cuentas por Cobrar");
						bOk = false;
					}
					if (!Funcion.Permiso("114", cdsSeguridad))
					{
						errorProvider.SetError(cmbTipoFact, "No tiene acceso a Cuentas por Pagar");
						bOk = false;
					}
				}
				#endregion No se permite mirar todo cuando hay restriccciones en cxc o cxp
				switch ((int) cmbTipoFact.EditValue)
				{
					case 1:
						if (!Funcion.Permiso("113", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Cuentas por Cobrar");
							bOk = false;
						}
						break;
					case 4:
						if (!Funcion.Permiso("114", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Cuentas por Pagar");
							bOk = false;
						}
						break;
					case 99:
						if (!Funcion.Permiso("113", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Cuentas por Pagar");
							bOk = false;
						}
						break;


					default:
							if (chkFiltroTipo.Checked == true)break;
							errorProvider.SetError(cmbTipoFact, "Solo son aplicables Opciones de Compra y Venta");
							bOk = false;
						break;
				}
			}
			#endregion Fin de Cuentas por Cobrar y Pagar
						
			#region Filtro de Abonos a Cobros y Pagos
			if ((int) rdgTipo.EditValue == 2)
			{
				switch ((int) cmbTipoFact.EditValue)
				{
					case 1:
						if (!Funcion.Permiso("117", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Abonos Recibidos");
							bOk = false;
						}
						break;
					case 4:
						if (!Funcion.Permiso("118", cdsSeguridad))
						{
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Abonos Realizados");
							bOk = false;
						}
						break;
					default:
						errorProvider.SetError(cmbTipoFact, "Solo son aplicables Opciones de Compra y Venta");
						bOk = false;
						break;
				}
			}
			#endregion Fin de Cuentas por Cobrar y Pagar

			return bOk;
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Seguridad()) return;
			if (!Verifica()) return;
			DateTime dtIni = DateTime.Now;
			string[] stFiltro1 =  new string[3];
			stFiltro1 = Filtro();
			string stFiltro = stFiltro1[0];
			string stFiltroRep = cmbReporte.ActiveRow.Cells["Filtro"].Value.ToString();
			if (stFiltroRep != null && stFiltroRep.Trim().Length > 0) 
				stFiltro += " And " + stFiltroRep;

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			if ((int) rdgTipo.EditValue == 1)
			{
				DateTime dtFechaFin = (DateTime) this.dtFechaHasta.EditValue;
				paramField.ParameterFieldName = "FechaFin";
				discreteVal.Value = dtFechaFin.ToString("yyyy-MM-dd");
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);
			}
			string stTitulo = GeneraTitulo() + stFiltro1[1];

			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'GEUR'"))
			{
				try
				{
					this.cdsSeteoRepF.Update();
				}
				catch
				{
				}
			}
			dr = this.cdsSeteoRepF.TableViews["SeteoRepFact"].Rows[0];
			#region Saldo CxC del Periodo
			string stReporte = this.cmbReporte.Text.Trim() + ".rpt";

			#region SaldoPeriodo
			if (stReporte.ToUpper().StartsWith("SALDOSPERIODO"))
			{
				DateTime dtFechaIni = (DateTime) this.dtFechaDesde.EditValue;
				DateTime dtFechaFin = (DateTime) this.dtFechaHasta.EditValue;
				stTitulo = GeneraTitulo()+ " Desde " + dtFechaIni.ToString("dd/MMM/yyyy")
					+ stFiltro1[1];;
				paramFields.Clear();
				ParameterField paramField1 = new ParameterField ();
				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "@FechaIni";
				discreteVal1.Value = dtFechaIni.ToString("yyyy-MM-dd HH:mm:ss");
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);
		
				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@FechaFin";
				discreteVal2.Value = dtFechaFin.ToString("yyyy-MM-dd HH:mm:ss");
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				ParameterField paramField3 = new ParameterField ();
				ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
				paramField3.ParameterFieldName = "@Signo";
				int iSigno = 1;
				if ((int) cmbTipoFact.EditValue == 1) iSigno = -1;
				discreteVal3.Value = iSigno;
				paramField3.CurrentValues.Add (discreteVal3);
				paramFields.Add (paramField3);
				
				string stFiltro3 = stFiltro1[2];
				ParameterField paramField4 = new ParameterField ();
				ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
				paramField4.ParameterFieldName = "@Filtro";
				discreteVal4.Value = stFiltro3;
				paramField4.CurrentValues.Add (discreteVal4);
				paramFields.Add (paramField4);
				stFiltro = "";
			}
			#endregion SaldoPeriodo
			
			#region Saldo por Proyecto
			if (stReporte.ToUpper().StartsWith("SALDOPROY"))
			{
				DateTime dtFechaIni = (DateTime) this.dtFechaDesde.EditValue;
				DateTime dtFechaFin = (DateTime) this.dtFechaHasta.EditValue;
				stTitulo = GeneraTitulo()+ " Desde " + dtFechaIni.ToString("dd/MMM/yyyy")
					+ stFiltro1[1];
				paramFields.Clear();
				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@FechaFin";
				discreteVal2.Value = dtFechaFin.ToString("yyyy-MM-dd HH:mm:ss");
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				int idClienteProy = 0;
				int idProyectoProy = 0;
				if (!chkCentro.Checked) idProyectoProy = (int) cmbCentro.EditValue;

				ParameterField paramField3 = new ParameterField ();
				ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
				if (stReporte.ToUpper().StartsWith("SALDOPROYSUB"))
				{
					paramField3.ParameterFieldName = "@idSubProyecto";
				}
				else
				{
					if (!chkCliente.Checked) idClienteProy = (int) cmbCliente.EditValue;
					paramField3.ParameterFieldName = "@idCliente";
				}
				discreteVal3.Value = idClienteProy;
				paramField3.CurrentValues.Add (discreteVal3);
				paramFields.Add (paramField3);
				
				string stFiltro3 = stFiltro1[2];
				ParameterField paramField4 = new ParameterField ();
				ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
				paramField4.ParameterFieldName = "@idProyecto";
				discreteVal4.Value = idProyectoProy;
				paramField4.CurrentValues.Add (discreteVal4);
				paramFields.Add (paramField4);
				stFiltro = "";
			}
			#endregion Saldo por proyecto

			#region Egresos x Centro
			if (stReporte.ToUpper().StartsWith("EGRESOSXCENTRO"))
			{
				
				DateTime dtFechaIni  = (DateTime) this.dtFechaDesde.EditValue;
				DateTime dtFechaFin  = (DateTime) this.dtFechaHasta.EditValue;
				DateTime dtFechaIniD = dtFechaIni.AddDays(1);

				string stExec1="EXEC CREAEGRESOCENTROCOSTO '" + dtFechaIni.ToString("yyyyMMdd")  
					+ "','" + dtFechaFin.ToString("yyyyMMdd") + "','" + dtFechaIniD.ToString("yyyyMMdd") +"'";
					
				Funcion.EjecutaSQL(cdsCiudad,stExec1); 

				stTitulo = GeneraTitulo()+ " Desde " + dtFechaIni.ToString("dd/MMM/yyyy")
					+ stFiltro1[1];;
				paramFields.Clear();
				ParameterField paramField1 = new ParameterField ();
				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "@FechaIni";
				discreteVal1.Value = dtFechaIni.ToString("yyyy-MM-dd");
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);
		
				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@FechaFin";
				discreteVal2.Value = dtFechaFin.ToString("yyyy-MM-dd");
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);
				stFiltro = "";
			}
			#endregion Egreso por Centro

			#endregion Saldo CxC del Periodo

			this.Cursor = Cursors.WaitCursor;

			if ((int) cmbTipoFact.EditValue == 7) 
			{
				stFiltro = stFiltro.Replace("And {Compra.idTipoFactura} = 7", "");
				stFiltro = stFiltro.Replace("Compra", "OrdenProduccion");
			}
			if (!this.chkRango.Checked)
			{
				stFiltro += " And {Compra.Numero} >= '" + this.txtFacDesde.Text+"'" ;
				stFiltro += " And {Compra.Numero} <= '" + this.txtFacHasta.Text +"'";
			}

			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(stTitulo);
			if (chkResumen.Checked) miReporte.Resumen();
			if ((int) rdgTipo.EditValue == 1)
			{
					miReporte.crVista.ParameterFieldInfo = paramFields;
			}
			miReporte.Show();
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			btAceptar.ToolTip = "Tiempo de ultimo reporte : " + ts.TotalSeconds.ToString("n2");

			this.Cursor = Cursors.Default;
		}

		private void cmbFechas_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (dr == null) return;
			dr["FiltroFecha"] = (int) this.cmbFechas.SelectedIndex;
			int iTipo = (int) this.cmbFechas.SelectedIndex;
			DateTime dtHoy = DateTime.Today;
			int iDia = dtHoy.Day;
			DateTime dtMes1 = dtHoy.AddDays(-iDia+1);
			DateTime dtDesde = (DateTime) dr["FechaIni"];
			DateTime dtHasta = (DateTime) dr["FechaFin"];
			switch(iTipo)
			{
				case 0: // Hasta Hoy
					dtHasta = DateTime.Today;
					break;
				case 1: // Solo Hoy
					dtHasta = DateTime.Today;
					dtDesde = DateTime.Today;
					break;
				case 2: // Este Mes
					dtDesde = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					dtHasta = dtMes1.AddDays(-1); 
					break;
				case 3: // El mes pasado
					dtMes1 = dtMes1.AddMonths(-1);
					dtDesde = dtMes1;
					dtMes1 = dtMes1.AddMonths(1);
					dtHasta = dtMes1.AddDays(-1); 
					break;
				case 4: // Este Año
					dtDesde = DateTime.Parse("01/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("12/31/" 
						+ DateTime.Today.Year.ToString(),us);
					break;
				case 5: // El Año pasado
					dtDesde = DateTime.Parse("01/01/" 
						+ DateTime.Today.AddYears(-1).Year.ToString(), us);
					dtHasta = DateTime.Parse("12/31/" 
						+ DateTime.Today.AddYears(-1).Year.ToString(), us);
					break;
				case 6: // El Periodo Ingresado
					dtDesde = Funcion.FechaCompra(this.cdsCliente, 0);
					dtHasta = Funcion.FechaCompra(this.cdsCliente, 1);
					break;
				case 7: // El Primer Trimestre
					dtDesde = DateTime.Parse("01/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("03/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 8: // El segundo Trimestre
					dtDesde = DateTime.Parse("04/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("06/30/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 9: // El tercer Trimestre
					dtDesde = DateTime.Parse("07/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("09/30/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 10: // El cuarto Trimestre
					dtDesde = DateTime.Parse("10/01/" 
						+ DateTime.Today.Year.ToString(), us);
					dtHasta = DateTime.Parse("12/31/" 
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 11: // Perido Digitado
					break;
			}
			dr["FechaIni"] = dtDesde;
			dr["FechaFin"] = dtHasta;
//			this.cdsSeteoRepF.Update();
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCliente.Enabled = !this.chkCliente.Checked;
		}

		private void chkGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkGrupo.CheckState == System.Windows.Forms.CheckState.Checked)
			{
				cmbGrupo.Enabled = false;
			}
			else
			{
				cmbGrupo.Enabled = true;
				cmbGrupo.Select();
			}
		}

		private void chkCentro_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkCentro.CheckState != System.Windows.Forms.CheckState.Indeterminate)
				cmbCentro.Enabled = !chkCentro.Checked;
		}

		private void chkComprobante_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkComprobante.CheckState != System.Windows.Forms.CheckState.Indeterminate)
				cmbComprobante.Enabled = !chkComprobante.Checked;
		}

		private void chkCiudad_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkCiudad.CheckState != System.Windows.Forms.CheckState.Indeterminate)
				cmbCiudad.Enabled = !chkCiudad.Checked;
		}

		private void chkTipoTrans_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkTipoTrans.CheckState != System.Windows.Forms.CheckState.Indeterminate)
				cmbTipoFact.Enabled = !chkTipoTrans.Checked;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}


		C1.Data.C1DataRow dr;
		C1.Data.C1DataRow drSeteoF;
		private void RepFacturacion_Load(object sender, System.EventArgs e)
		{
			#region Tres Estados en Check
			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'FTMTE'")) 				
			{
				chkCentro.Properties.AllowGrayed = true;
				chkCiudad.Properties.AllowGrayed = true;
				chkCliente.Properties.AllowGrayed = true;
				chkComprobante.Properties.AllowGrayed = true;
				chkCredTribut.Properties.AllowGrayed = true;
				chkFormaPago.Properties.AllowGrayed = true;
				chkGrupo.Properties.AllowGrayed = true;
				chkOtro.Properties.AllowGrayed = true;
				chkTipoTrans.Properties.AllowGrayed = true;
				chkVendedor.Properties.AllowGrayed = true;
			}
			#endregion Tres Estados en Check
			
			string stAudita = "Exec AuditaCrear 53, 6, 'RepFacturacion'";
			Funcion.EjecutaSQL(cdsCiudad, stAudita, true);
			#region activar impresion por lotes por numeros

			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'FIMPL'";
			if (Funcion.bEjecutaSQL(cdsCiudad, stActivado))
			{
				chkRango.Visible = true;
				txtFacDesde.Visible = true;
				txtFacHasta.Visible = true;
			}

			#endregion activar impresion por lotes por numeros

			Cursor = Cursors.Default;
			#region Cargar Combos
			cmbVendedor.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoRepF,
				"Select Nombre, idPersonal From Personal Where Vendedor = 1");
			cmbCredTribut.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoRepF,
				"Select idCredTributario, CredTributario From CompraCredTribut");
			string stExec = "Exec ClienteCargaNombre 0";
			cmbCliente.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoRepF, stExec);

			#endregion Cargar Combos

			#region Visibilidad de Combos
			chkCentro_CheckedChanged(this, e);
			chkCiudad_CheckedChanged(this, e);
			chkCliente_CheckedChanged(this, e);
			chkComprobante_CheckedChanged(this, e);
			chkCredTribut_CheckedChanged(this, e);
			chkFormaPago_CheckedChanged(this, e);
			chkGrupo_CheckedChanged(this, e);
			chkOtro_CheckedChanged(this, e);
			chkTipoTrans_CheckedChanged(this, e);
			chkVendedor_CheckedChanged(this, e);
			#endregion Visibilidad de Combos
			
			dr = cdsSeteoRepF.TableViews["SeteoRepFact"].Rows[0];
			drSeteoF = cdsSeteoRepF.TableViews["SeteoF"].Rows[0];
			if (dr["FiltroFecha"] == System.DBNull.Value)
				cmbFechas.SelectedIndex = 0;
			else
        cmbFechas.SelectedIndex = (int) dr["FiltroFecha"];

//			#region Bodega por Nombre
//			if (drSeteoF["VerBodegaNombre"] != DBNull.Value && (bool) drSeteoF["VerBodegaNombre"] == true)
//			{
//				spnBodega.Width = 0;
//				udsBodegas.Rows.Clear();
//				string stExecBodega = "Exec BodegaUsuario 0"	;//+ IdTipoFactura.ToString();
//				SqlDataReader drBod = Funcion.rEscalarSQL(cdsCliente, stExecBodega, true);
//				int iNumFilas = 0;
//				while(drBod.Read())
//				{
//					udsBodegas.Rows.Add();
//					udsBodegas.Rows[iNumFilas]["Nombre"]=drBod.GetString(0);
//					udsBodegas.Rows[iNumFilas]["Bodega"]=drBod.GetInt32(1);
//					iNumFilas++;
//				}
//			}
//			else
//			{
//				cmbBodega.Width = 0;
//			}
//			#endregion Bodega por Nombre
		}

		private void RepFacturacion_Activated(object sender, System.EventArgs e)
		{
			this.cmbTipoFact.Focus();
		}

		private void chkFormaPago_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbFormaPago.Enabled = !this.chkFormaPago.Checked;
		}

		private void cdsTipoFact_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipoFact.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsGrupoCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupoCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCiudad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCiudad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsComprobante_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsComprobante.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFormaPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsSeteoRepF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoRepF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsReporte_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
//			Funcion miFuncion = new Funcion();
//			this.cdsReporte.Schema.Connections[0].ConnectionString = 
//				miFuncion.AccesoBase(MenuLatinium.stDirInicio);
		}

		private void rdgTipo_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			int IdGrupo = (int) this.rdgTipo.EditValue + 1;
			string stFiltro = "Tipo = 0 And idGrupo = " + IdGrupo.ToString();
			this.cdvReporte.RowFilter = stFiltro;
			DataRow[] drRep = this.cdsReporte.StorageDataSet.Tables["Reporte"].Select(stFiltro);
			if (drRep.Length > 0)
				dr["TipoReporte"] = drRep[0]["idReporte"];
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btCierreCaja_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFechaIni = (DateTime) dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) dtFechaHasta.EditValue;
			#region Cuardre de Caja con Procedimiento Almacenado
			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'FPVCCUA'"))
			{
				CierreCaja miCaja = new CierreCaja();
				miCaja.MdiParent = this.MdiParent;
				miCaja.Show();
				Cursor = Cursors.Default;
				return;
			}
			#endregion Cuardre de Caja con Procedimiento Almacenado

			string stSelect = "Exec CreaCierreCaja '" + dtFechaIni.ToString("yyyyMMdd") + "', '" + dtFechaFin.ToString("yyyyMMdd") + "'";
			if (chkTipoTrans.Checked) stSelect += ", 0";
			else stSelect += ", " + cmbTipoFact.EditValue.ToString();
			Funcion.EjecutaSQL(cdsCliente, stSelect, false);
			string stFiltro = "";
			string stTitulo = "";
			if (this.chkCentro.Checked == false)
			{
				stFiltro = " {TmpCierre.idProyecto} = " + cmbCentro.EditValue.ToString();
				stTitulo += "Centro de Costo " + cmbCentro.Text.Trim() + " ";
			}
			Reporte miRep = new Reporte("CierreCaja", stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo(stTitulo);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void chkVendedor_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbVendedor.Enabled = !this.chkVendedor.Checked;		
		}

		private void cmbCliente_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCliente.Properties.Columns.Clear();
			this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCliente.Properties.DataSource = this.cdvCodigoCliente;
			this.cmbCliente.Properties.DisplayMember = "Codigo";
		}

		private void RepFacturacion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepFacturacion'";
			Funcion.EjecutaSQL(cdsCiudad, stAudita, true);
		}

		private void chkCredTribut_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCredTribut.Enabled = !chkCredTribut.Checked;
			if (!chkCredTribut.Checked) cmbCredTribut.Select();
		}

		private void cdsCompraTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompraTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void chkCredTribut_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCredTribut.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCredTribut.Enabled = true;
			cmbCredTribut.Select();
		}

		private void chkComprobante_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkComprobante.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbComprobante.Enabled = true;
			cmbComprobante.Select();		
		}

		private void chkOtro_CheckedChanged(object sender, System.EventArgs e)
		{
			this.txtOtro.Enabled = !this.chkOtro.Checked;
		}

		private void btCierreCaja_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			CierreCaja miCierre = new CierreCaja();
			miCierre.MdiParent = this.MdiParent;
			miCierre.Show();
		}

		private void chkRango_CheckedChanged(object sender, System.EventArgs e)
		{
			this.txtFacDesde.Enabled = !this.chkRango.Checked;
			this.txtFacHasta.Enabled = !this.chkRango.Checked;

		}

		private void chkCliente_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCliente.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCliente.Enabled = true;
			cmbCliente.Select();
		}

		private void chkGrupo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkGrupo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbGrupo.Enabled = true;
			cmbGrupo.Select();
		}

		private void chkCentro_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCentro.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCentro.Enabled = true;
			cmbCentro.Select();
		}

		private void chkCiudad_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCiudad.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCiudad.Enabled = true;
			cmbCiudad.Select();
		}

		private void chkFormaPago_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkFormaPago.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbFormaPago.Enabled = true;
			cmbFormaPago.Select();
		}

		private void chkVendedor_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkVendedor.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbVendedor.Enabled = true;
			cmbVendedor.Select();
		}

		private void dtFechaDesde_EditValueChanged(object sender, System.EventArgs e)
		{
		
		}

		

	}
}
