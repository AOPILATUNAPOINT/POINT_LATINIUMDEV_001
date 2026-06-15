using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepFacturacion.
	/// </summary>
	public class RepArticulos : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsTipoFact;
		private C1.Data.C1DataView cdvTipoFact;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.LookUpEdit cmbTipoFact;
		private DevExpress.XtraEditors.ComboBoxEdit cmbFechas;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.DateEdit dtFechaHasta;
		private DevExpress.XtraEditors.DateEdit dtFechaDesde;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private DevExpress.XtraEditors.LookUpEdit cmbCliente;
		private DevExpress.XtraEditors.LookUpEdit cmbGrupo;
		private DevExpress.XtraEditors.CheckEdit chkGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbUsuario;
		private DevExpress.XtraEditors.CheckEdit chkUsuario;
		private DevExpress.XtraEditors.LookUpEdit cmbCentro;
		private DevExpress.XtraEditors.CheckEdit chkCentro;
		private DevExpress.XtraEditors.LookUpEdit cmbSubCentro;
		private DevExpress.XtraEditors.CheckEdit chkSubCentro;
		private DevExpress.XtraEditors.RadioGroup rdgFecha;
		private DevExpress.XtraEditors.LookUpEdit cmbComprobante;
		private DevExpress.XtraEditors.CheckEdit chkComprobante;
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
		private DevExpress.XtraEditors.LookUpEdit cmbSector;
		private DevExpress.XtraEditors.CheckEdit chkSector;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.SpinEdit spnBodega;
		private DevExpress.XtraEditors.CheckEdit chkBodega;
		private DevExpress.XtraEditors.CheckEdit chkResumen;
		private DevExpress.XtraEditors.CheckEdit chkAnulado;
		private C1.Data.C1DataSet cdsCiudad;
		private C1.Data.C1DataView cdvCiudad;
		private C1.Data.C1DataSet cdsComprobante;
		private C1.Data.C1DataView cdvComprobante;
		private C1.Data.C1DataView cdvSector;
		private C1.Data.C1DataView cdvReporte;
		private DevExpress.XtraEditors.LookUpEdit cmbGrupoArt;
		private DevExpress.XtraEditors.CheckEdit chkGrupoArt;
		private DevExpress.XtraEditors.CheckEdit chkArticulo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsGrupoArt;
		private C1.Data.C1DataView cdvGrupoArt;
		private DevExpress.XtraEditors.LookUpEdit cmbTipo;
		private C1.Data.C1DataView cdvTipo;
		private C1.Data.C1DataSet cdsTipo;
		private C1.Data.C1DataView cdvArtCodigo;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private DevExpress.XtraEditors.CheckEdit chkModelo;
		private DevExpress.XtraEditors.CheckEdit chkSubGrupo;
		private DevExpress.XtraEditors.CheckEdit chkClase;
		private DevExpress.XtraEditors.CheckEdit chkColor;
		private DevExpress.XtraEditors.CheckEdit chkMarca;
		private C1.Data.C1DataView cdvClienteCod;
		private DevExpress.XtraEditors.LookUpEdit cmbArtCod;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraEditors.CheckEdit chkSRI;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbImpuesto;
		private DevExpress.XtraEditors.CheckEdit chkCredTribut;
		private C1.Data.C1DataSet cdsCompraTabla;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsBodegas;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private DevExpress.XtraEditors.CheckEdit chkNeto;
		private DevExpress.XtraEditors.CheckEdit chkTipo;
		private DevExpress.XtraEditors.LookUpEdit cmbSubGrupo;
		private DevExpress.XtraEditors.LookUpEdit cmbModelo;
		private DevExpress.XtraEditors.LookUpEdit cmbColor;
		private DevExpress.XtraEditors.LookUpEdit cmbClase;
		private DevExpress.XtraEditors.LookUpEdit cmbCredTribut;
		private DevExpress.XtraEditors.LookUpEdit cmbSRI;
		private DevExpress.XtraEditors.LookUpEdit cmbMarca;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepArticulos()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			this.cmbTipoFact = new DevExpress.XtraEditors.LookUpEdit();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdvTipoFact = new C1.Data.C1DataView();
			this.cdsTipoFact = new C1.Data.C1DataSet();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbFechas = new DevExpress.XtraEditors.ComboBoxEdit();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFechaHasta = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaDesde = new DevExpress.XtraEditors.DateEdit();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.cmbCliente = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvGrupoCliente = new C1.Data.C1DataView();
			this.cdsGrupoCliente = new C1.Data.C1DataSet();
			this.chkGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbUsuario = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.chkUsuario = new DevExpress.XtraEditors.CheckEdit();
			this.cmbCentro = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.chkCentro = new DevExpress.XtraEditors.CheckEdit();
			this.cmbSubCentro = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.chkSubCentro = new DevExpress.XtraEditors.CheckEdit();
			this.rdgFecha = new DevExpress.XtraEditors.RadioGroup();
			this.cmbComprobante = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvComprobante = new C1.Data.C1DataView();
			this.cdsComprobante = new C1.Data.C1DataSet();
			this.chkComprobante = new DevExpress.XtraEditors.CheckEdit();
			this.chkTipoTrans = new DevExpress.XtraEditors.CheckEdit();
			this.chkResumen = new DevExpress.XtraEditors.CheckEdit();
			this.chkAnulado = new DevExpress.XtraEditors.CheckEdit();
			this.cmbCiudad = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCiudad = new C1.Data.C1DataView();
			this.cdsCiudad = new C1.Data.C1DataSet();
			this.chkCiudad = new DevExpress.XtraEditors.CheckEdit();
			this.cmbSector = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvSector = new C1.Data.C1DataView();
			this.chkSector = new DevExpress.XtraEditors.CheckEdit();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.spnBodega = new DevExpress.XtraEditors.SpinEdit();
			this.chkBodega = new DevExpress.XtraEditors.CheckEdit();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.cmbGrupoArt = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvGrupoArt = new C1.Data.C1DataView();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.chkGrupoArt = new DevExpress.XtraEditors.CheckEdit();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.chkArticulo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbTipo = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvTipo = new C1.Data.C1DataView();
			this.cdsTipo = new C1.Data.C1DataSet();
			this.cmbArtCod = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkModelo = new DevExpress.XtraEditors.CheckEdit();
			this.chkSubGrupo = new DevExpress.XtraEditors.CheckEdit();
			this.chkClase = new DevExpress.XtraEditors.CheckEdit();
			this.chkColor = new DevExpress.XtraEditors.CheckEdit();
			this.chkMarca = new DevExpress.XtraEditors.CheckEdit();
			this.cdvClienteCod = new C1.Data.C1DataView();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.chkSRI = new DevExpress.XtraEditors.CheckEdit();
			this.cmbImpuesto = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.chkCredTribut = new DevExpress.XtraEditors.CheckEdit();
			this.udsBodegas = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkNeto = new DevExpress.XtraEditors.CheckEdit();
			this.chkTipo = new DevExpress.XtraEditors.CheckEdit();
			this.cmbSubGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbModelo = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbColor = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbClase = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbCredTribut = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbSRI = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbMarca = new DevExpress.XtraEditors.LookUpEdit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkUsuario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCentro.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoTrans.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumen.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnulado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCiudad.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSector.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBodega.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoArt.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupoArt.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkArticulo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtCod.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkModelo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkClase.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkColor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMarca.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClienteCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSRI.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImpuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCredTribut.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNeto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSRI.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbTipoFact
			// 
			this.cmbTipoFact.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Transaccion"));
			this.cmbTipoFact.Location = new System.Drawing.Point(384, 17);
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
			this.cmbTipoFact.Properties.PopupWidth = 300;
			this.cmbTipoFact.Properties.ShowFooter = false;
			this.cmbTipoFact.Properties.ShowHeader = false;
			this.cmbTipoFact.Properties.ValueMember = "idTipoFactura";
			this.cmbTipoFact.Size = new System.Drawing.Size(176, 20);
			this.cmbTipoFact.TabIndex = 1;
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// cdvTipoFact
			// 
			this.cdvTipoFact.BindingContextCtrl = this;
			this.cdvTipoFact.DataSet = this.cdsTipoFact;
			this.cdvTipoFact.RowFilter = "Grupo = 1 Or Grupo = 2";
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
			this.btAceptar.Location = new System.Drawing.Point(48, 362);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 2;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(48, 414);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 3;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbFechas
			// 
			this.cmbFechas.EditValue = "Hasta Hoy";
			this.cmbFechas.Location = new System.Drawing.Point(16, 250);
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
			this.cmbFechas.Size = new System.Drawing.Size(160, 20);
			this.cmbFechas.TabIndex = 52;
			this.cmbFechas.SelectedIndexChanged += new System.EventHandler(this.cmbFechas_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 302);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 51;
			this.label2.Text = "Hasta:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 276);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 50;
			this.label1.Text = "Desde:";
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.FechaFin"));
			this.dtFechaHasta.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaHasta.Location = new System.Drawing.Point(80, 302);
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
			this.dtFechaHasta.TabIndex = 49;
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.FechaIni"));
			this.dtFechaDesde.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFechaDesde.Location = new System.Drawing.Point(80, 276);
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
			this.dtFechaDesde.TabIndex = 48;
			// 
			// chkCliente
			// 
			this.chkCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bCliente"));
			this.chkCliente.EditValue = true;
			this.chkCliente.Location = new System.Drawing.Point(216, 42);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.Caption = "Todo Cliente";
			this.chkCliente.Size = new System.Drawing.Size(96, 19);
			this.chkCliente.TabIndex = 53;
			this.chkCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkCliente_MouseUp);
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// cmbCliente
			// 
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Cliente"));
			this.cmbCliente.Location = new System.Drawing.Point(384, 42);
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
			this.cmbCliente.Properties.PopupWidth = 400;
			this.cmbCliente.Properties.ShowHeader = false;
			this.cmbCliente.Properties.ValueMember = "idCliente";
			this.cmbCliente.Size = new System.Drawing.Size(176, 20);
			this.cmbCliente.TabIndex = 54;
			this.cmbCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCliente_MouseUp);
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
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Grupo"));
			this.cmbGrupo.Location = new System.Drawing.Point(384, 67);
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
			this.cmbGrupo.Properties.PopupWidth = 400;
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
			this.chkGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bGrupo"));
			this.chkGrupo.EditValue = true;
			this.chkGrupo.Location = new System.Drawing.Point(216, 67);
			this.chkGrupo.Name = "chkGrupo";
			// 
			// chkGrupo.Properties
			// 
			this.chkGrupo.Properties.Caption = "Todo Grupo Cliente";
			this.chkGrupo.Size = new System.Drawing.Size(136, 19);
			this.chkGrupo.TabIndex = 55;
			this.chkGrupo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkGrupo_MouseUp);
			this.chkGrupo.CheckedChanged += new System.EventHandler(this.chkGrupo_CheckedChanged);
			// 
			// cmbUsuario
			// 
			this.cmbUsuario.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Vendedor"));
			this.cmbUsuario.Location = new System.Drawing.Point(384, 92);
			this.cmbUsuario.Name = "cmbUsuario";
			// 
			// cmbUsuario.Properties
			// 
			this.cmbUsuario.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbUsuario.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbUsuario.Properties.DataSource = this.cdvPersonal;
			this.cmbUsuario.Properties.DisplayMember = "Nombre";
			this.cmbUsuario.Properties.NullText = "";
			this.cmbUsuario.Properties.PopupWidth = 400;
			this.cmbUsuario.Properties.ShowFooter = false;
			this.cmbUsuario.Properties.ShowHeader = false;
			this.cmbUsuario.Properties.ValueMember = "idPersonal";
			this.cmbUsuario.Size = new System.Drawing.Size(176, 20);
			this.cmbUsuario.TabIndex = 58;
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
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
			// chkUsuario
			// 
			this.chkUsuario.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bVendedor"));
			this.chkUsuario.EditValue = true;
			this.chkUsuario.Location = new System.Drawing.Point(216, 92);
			this.chkUsuario.Name = "chkUsuario";
			// 
			// chkUsuario.Properties
			// 
			this.chkUsuario.Properties.Caption = "Todo Vendedor";
			this.chkUsuario.Size = new System.Drawing.Size(112, 19);
			this.chkUsuario.TabIndex = 57;
			this.chkUsuario.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkUsuario_MouseUp);
			this.chkUsuario.CheckedChanged += new System.EventHandler(this.chkUsuario_CheckedChanged);
			// 
			// cmbCentro
			// 
			this.cmbCentro.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Centro"));
			this.cmbCentro.Location = new System.Drawing.Point(384, 117);
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
			this.cmbCentro.Properties.PopupWidth = 400;
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
			this.chkCentro.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bCentro"));
			this.chkCentro.EditValue = true;
			this.chkCentro.Location = new System.Drawing.Point(216, 117);
			this.chkCentro.Name = "chkCentro";
			// 
			// chkCentro.Properties
			// 
			this.chkCentro.Properties.Caption = "Todo Centro Costo";
			this.chkCentro.Size = new System.Drawing.Size(120, 19);
			this.chkCentro.TabIndex = 59;
			this.chkCentro.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkCentro_MouseUp);
			this.chkCentro.CheckedChanged += new System.EventHandler(this.chkCentro_CheckedChanged);
			// 
			// cmbSubCentro
			// 
			this.cmbSubCentro.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.SubCentro"));
			this.cmbSubCentro.Location = new System.Drawing.Point(384, 142);
			this.cmbSubCentro.Name = "cmbSubCentro";
			// 
			// cmbSubCentro.Properties
			// 
			this.cmbSubCentro.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSubCentro.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSubCentro.Properties.DataSource = this.cdvSubProyecto;
			this.cmbSubCentro.Properties.DisplayMember = "Nombre";
			this.cmbSubCentro.Properties.NullText = "";
			this.cmbSubCentro.Properties.PopupWidth = 400;
			this.cmbSubCentro.Properties.ShowFooter = false;
			this.cmbSubCentro.Properties.ShowHeader = false;
			this.cmbSubCentro.Properties.ValueMember = "idSubProyecto";
			this.cmbSubCentro.Size = new System.Drawing.Size(176, 20);
			this.cmbSubCentro.TabIndex = 62;
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// chkSubCentro
			// 
			this.chkSubCentro.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bSubCentro"));
			this.chkSubCentro.EditValue = true;
			this.chkSubCentro.Location = new System.Drawing.Point(216, 142);
			this.chkSubCentro.Name = "chkSubCentro";
			// 
			// chkSubCentro.Properties
			// 
			this.chkSubCentro.Properties.Caption = "Todo Proyecto";
			this.chkSubCentro.Size = new System.Drawing.Size(112, 19);
			this.chkSubCentro.TabIndex = 61;
			this.chkSubCentro.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkSubCentro_MouseUp);
			this.chkSubCentro.CheckedChanged += new System.EventHandler(this.chkSubCentro_CheckedChanged);
			// 
			// rdgFecha
			// 
			this.rdgFecha.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.TipoFecha"));
			this.rdgFecha.EditValue = 0;
			this.rdgFecha.Location = new System.Drawing.Point(48, 164);
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
			this.rdgFecha.Size = new System.Drawing.Size(96, 69);
			this.rdgFecha.TabIndex = 63;
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Comprobante"));
			this.cmbComprobante.Location = new System.Drawing.Point(384, 167);
			this.cmbComprobante.Name = "cmbComprobante";
			// 
			// cmbComprobante.Properties
			// 
			this.cmbComprobante.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbComprobante.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Comprobante", "Comprobante", 78, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbComprobante.Properties.DataSource = this.cdvComprobante;
			this.cmbComprobante.Properties.DisplayMember = "Comprobante";
			this.cmbComprobante.Properties.NullText = "";
			this.cmbComprobante.Properties.PopupWidth = 400;
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
			// chkComprobante
			// 
			this.chkComprobante.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bComprobante"));
			this.chkComprobante.EditValue = true;
			this.chkComprobante.Location = new System.Drawing.Point(216, 167);
			this.chkComprobante.Name = "chkComprobante";
			// 
			// chkComprobante.Properties
			// 
			this.chkComprobante.Properties.Caption = "Todo Comprobante";
			this.chkComprobante.Size = new System.Drawing.Size(128, 19);
			this.chkComprobante.TabIndex = 64;
			this.chkComprobante.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkComprobante_MouseUp);
			this.chkComprobante.CheckedChanged += new System.EventHandler(this.chkComprobante_CheckedChanged);
			// 
			// chkTipoTrans
			// 
			this.chkTipoTrans.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bTransaccion"));
			this.chkTipoTrans.Location = new System.Drawing.Point(216, 17);
			this.chkTipoTrans.Name = "chkTipoTrans";
			// 
			// chkTipoTrans.Properties
			// 
			this.chkTipoTrans.Properties.Caption = "Toda Transacción";
			this.chkTipoTrans.Size = new System.Drawing.Size(144, 19);
			this.chkTipoTrans.TabIndex = 66;
			this.chkTipoTrans.CheckedChanged += new System.EventHandler(this.chkTipoTrans_CheckedChanged);
			// 
			// chkResumen
			// 
			this.chkResumen.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Resumen"));
			this.chkResumen.Location = new System.Drawing.Point(24, 17);
			this.chkResumen.Name = "chkResumen";
			// 
			// chkResumen.Properties
			// 
			this.chkResumen.Properties.Caption = "Resumen";
			this.chkResumen.Size = new System.Drawing.Size(75, 19);
			this.chkResumen.TabIndex = 67;
			// 
			// chkAnulado
			// 
			this.chkAnulado.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Anulados"));
			this.chkAnulado.Location = new System.Drawing.Point(24, 43);
			this.chkAnulado.Name = "chkAnulado";
			// 
			// chkAnulado.Properties
			// 
			this.chkAnulado.Properties.Caption = "Mostrar Anulados";
			this.chkAnulado.Size = new System.Drawing.Size(128, 19);
			this.chkAnulado.TabIndex = 68;
			// 
			// cmbCiudad
			// 
			this.cmbCiudad.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Ciudad"));
			this.cmbCiudad.Location = new System.Drawing.Point(384, 192);
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
			this.cmbCiudad.Properties.PopupWidth = 400;
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
			this.chkCiudad.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bCiudad"));
			this.chkCiudad.EditValue = true;
			this.chkCiudad.Location = new System.Drawing.Point(216, 192);
			this.chkCiudad.Name = "chkCiudad";
			// 
			// chkCiudad.Properties
			// 
			this.chkCiudad.Properties.Caption = "Toda Ciudad";
			this.chkCiudad.Size = new System.Drawing.Size(120, 19);
			this.chkCiudad.TabIndex = 69;
			this.chkCiudad.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkCiudad_MouseUp);
			this.chkCiudad.CheckedChanged += new System.EventHandler(this.chkCiudad_CheckedChanged);
			// 
			// cmbSector
			// 
			this.cmbSector.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Sector"));
			this.cmbSector.Location = new System.Drawing.Point(384, 217);
			this.cmbSector.Name = "cmbSector";
			// 
			// cmbSector.Properties
			// 
			this.cmbSector.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSector.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Sector", "Sector", 42, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSector.Properties.DataSource = this.cdvSector;
			this.cmbSector.Properties.DisplayMember = "Sector";
			this.cmbSector.Properties.NullText = "";
			this.cmbSector.Properties.PopupWidth = 400;
			this.cmbSector.Properties.ShowFooter = false;
			this.cmbSector.Properties.ShowHeader = false;
			this.cmbSector.Properties.ValueMember = "idSector";
			this.cmbSector.Size = new System.Drawing.Size(176, 20);
			this.cmbSector.TabIndex = 72;
			// 
			// cdvSector
			// 
			this.cdvSector.BindingContextCtrl = this;
			this.cdvSector.DataSet = this.cdsCiudad;
			this.cdvSector.TableName = "";
			this.cdvSector.TableViewName = "ClienteSector";
			// 
			// chkSector
			// 
			this.chkSector.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bSector"));
			this.chkSector.EditValue = true;
			this.chkSector.Location = new System.Drawing.Point(216, 217);
			this.chkSector.Name = "chkSector";
			// 
			// chkSector.Properties
			// 
			this.chkSector.Properties.Caption = "Todo Sector";
			this.chkSector.Size = new System.Drawing.Size(120, 19);
			this.chkSector.TabIndex = 71;
			this.chkSector.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkSector_MouseUp);
			this.chkSector.CheckedChanged += new System.EventHandler(this.chkSector_CheckedChanged);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// spnBodega
			// 
			this.spnBodega.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Bodega"));
			this.spnBodega.EditValue = new System.Decimal(new int[] {
																																1,
																																0,
																																0,
																																0});
			this.spnBodega.Location = new System.Drawing.Point(136, 95);
			this.spnBodega.Name = "spnBodega";
			// 
			// spnBodega.Properties
			// 
			this.spnBodega.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnBodega.Properties.IsFloatValue = false;
			this.spnBodega.Properties.Mask.EditMask = "N00";
			this.spnBodega.Properties.MaxValue = new System.Decimal(new int[] {
																																					9,
																																					0,
																																					0,
																																					0});
			this.spnBodega.Properties.MinValue = new System.Decimal(new int[] {
																																					1,
																																					0,
																																					0,
																																					0});
			this.spnBodega.Size = new System.Drawing.Size(48, 20);
			this.spnBodega.TabIndex = 73;
			// 
			// chkBodega
			// 
			this.chkBodega.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bBodega"));
			this.chkBodega.EditValue = true;
			this.chkBodega.Location = new System.Drawing.Point(24, 95);
			this.chkBodega.Name = "chkBodega";
			// 
			// chkBodega.Properties
			// 
			this.chkBodega.Properties.Caption = "Toda Bodega";
			this.chkBodega.Size = new System.Drawing.Size(96, 19);
			this.chkBodega.TabIndex = 74;
			this.chkBodega.CheckedChanged += new System.EventHandler(this.chkBodega_CheckedChanged);
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 1";
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
			// cmbGrupoArt
			// 
			this.cmbGrupoArt.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.GrupoArt"));
			this.cmbGrupoArt.Location = new System.Drawing.Point(384, 267);
			this.cmbGrupoArt.Name = "cmbGrupoArt";
			// 
			// cmbGrupoArt.Properties
			// 
			this.cmbGrupoArt.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbGrupoArt.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Grupo", "Grupo", 41, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbGrupoArt.Properties.DataSource = this.cdvGrupoArt;
			this.cmbGrupoArt.Properties.DisplayMember = "Grupo";
			this.cmbGrupoArt.Properties.NullText = "";
			this.cmbGrupoArt.Properties.PopupWidth = 400;
			this.cmbGrupoArt.Properties.ShowFooter = false;
			this.cmbGrupoArt.Properties.ShowHeader = false;
			this.cmbGrupoArt.Properties.ValueMember = "idGrupoArticulo";
			this.cmbGrupoArt.Size = new System.Drawing.Size(176, 20);
			this.cmbGrupoArt.TabIndex = 79;
			// 
			// cdvGrupoArt
			// 
			this.cdvGrupoArt.BindingContextCtrl = this;
			this.cdvGrupoArt.DataSet = this.cdsGrupoArt;
			this.cdvGrupoArt.TableName = "";
			this.cdvGrupoArt.TableViewName = "ArticuloGrupo";
			// 
			// cdsGrupoArt
			// 
			this.cdsGrupoArt.BindingContextCtrl = this;
			this.cdsGrupoArt.DataLibrary = "LibFacturacion";
			this.cdsGrupoArt.DataLibraryUrl = "";
			this.cdsGrupoArt.DataSetDef = "dsArticuloTabla";
			this.cdsGrupoArt.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoArt.Name = "cdsGrupoArt";
			this.cdsGrupoArt.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoArt.SchemaDef = null;
			this.cdsGrupoArt.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupoArt_BeforeFill);
			// 
			// chkGrupoArt
			// 
			this.chkGrupoArt.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bGrupoArt"));
			this.chkGrupoArt.EditValue = true;
			this.chkGrupoArt.Location = new System.Drawing.Point(216, 267);
			this.chkGrupoArt.Name = "chkGrupoArt";
			// 
			// chkGrupoArt.Properties
			// 
			this.chkGrupoArt.Properties.Caption = "Todo Grupo Art.";
			this.chkGrupoArt.Size = new System.Drawing.Size(120, 19);
			this.chkGrupoArt.TabIndex = 78;
			this.chkGrupoArt.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkGrupoArt_MouseUp);
			this.chkGrupoArt.CheckedChanged += new System.EventHandler(this.chkGrupoArt_CheckedChanged);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// chkArticulo
			// 
			this.chkArticulo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bArticulo"));
			this.chkArticulo.EditValue = true;
			this.chkArticulo.Location = new System.Drawing.Point(216, 242);
			this.chkArticulo.Name = "chkArticulo";
			// 
			// chkArticulo.Properties
			// 
			this.chkArticulo.Properties.Caption = "Todo Artículo";
			this.chkArticulo.Size = new System.Drawing.Size(120, 19);
			this.chkArticulo.TabIndex = 76;
			this.chkArticulo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkArticulo_MouseUp);
			this.chkArticulo.CheckedChanged += new System.EventHandler(this.chkArticulo_CheckedChanged);
			// 
			// cmbTipo
			// 
			this.cmbTipo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Tipo"));
			this.cmbTipo.Location = new System.Drawing.Point(384, 292);
			this.cmbTipo.Name = "cmbTipo";
			// 
			// cmbTipo.Properties
			// 
			this.cmbTipo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbTipo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Tipo", "Tipo", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbTipo.Properties.DataSource = this.cdvTipo;
			this.cmbTipo.Properties.DisplayMember = "Tipo";
			this.cmbTipo.Properties.NullText = "";
			this.cmbTipo.Properties.PopupWidth = 400;
			this.cmbTipo.Properties.ValueMember = "idTipoGrupo";
			this.cmbTipo.Size = new System.Drawing.Size(176, 20);
			this.cmbTipo.TabIndex = 80;
			// 
			// cdvTipo
			// 
			this.cdvTipo.BindingContextCtrl = this;
			this.cdvTipo.DataSet = this.cdsTipo;
			this.cdvTipo.TableName = "";
			this.cdvTipo.TableViewName = "ArticuloTipo";
			// 
			// cdsTipo
			// 
			this.cdsTipo.BindingContextCtrl = this;
			this.cdsTipo.DataLibrary = "LibFacturacion";
			this.cdsTipo.DataLibraryUrl = "";
			this.cdsTipo.DataSetDef = "dsArticuloTabla";
			this.cdsTipo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipo.Name = "cdsTipo";
			this.cdsTipo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipo.SchemaDef = null;
			this.cdsTipo.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipo_BeforeFill);
			// 
			// cmbArtCod
			// 
			this.cmbArtCod.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.Articulo"));
			this.cmbArtCod.Location = new System.Drawing.Point(384, 242);
			this.cmbArtCod.Name = "cmbArtCod";
			// 
			// cmbArtCod.Properties
			// 
			this.cmbArtCod.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbArtCod.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Artículo", 180, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 80, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbArtCod.Properties.DataSource = this.cdvArticulo;
			this.cmbArtCod.Properties.DisplayMember = "Articulo";
			this.cmbArtCod.Properties.NullText = "";
			this.cmbArtCod.Properties.PopupWidth = 400;
			this.cmbArtCod.Properties.ValueMember = "idArticulo";
			this.cmbArtCod.Size = new System.Drawing.Size(176, 20);
			this.cmbArtCod.TabIndex = 84;
			this.cmbArtCod.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbArtCod_MouseUp);
			// 
			// cdvArtCodigo
			// 
			this.cdvArtCodigo.BindingContextCtrl = this;
			this.cdvArtCodigo.DataSet = this.cdsArticulo;
			this.cdvArtCodigo.Sort = "Codigo";
			this.cdvArtCodigo.TableName = "";
			this.cdvArtCodigo.TableViewName = "Articulo";
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
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
			this.cmbReporte.Location = new System.Drawing.Point(24, 69);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(160, 21);
			this.cmbReporte.TabIndex = 85;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// chkModelo
			// 
			this.chkModelo.EditValue = true;
			this.chkModelo.Location = new System.Drawing.Point(216, 342);
			this.chkModelo.Name = "chkModelo";
			// 
			// chkModelo.Properties
			// 
			this.chkModelo.Properties.Caption = "Todo Modelo";
			this.chkModelo.Size = new System.Drawing.Size(120, 19);
			this.chkModelo.TabIndex = 88;
			this.chkModelo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkModelo_MouseUp);
			this.chkModelo.CheckedChanged += new System.EventHandler(this.chkModelo_CheckedChanged);
			// 
			// chkSubGrupo
			// 
			this.chkSubGrupo.EditValue = true;
			this.chkSubGrupo.Location = new System.Drawing.Point(216, 317);
			this.chkSubGrupo.Name = "chkSubGrupo";
			// 
			// chkSubGrupo.Properties
			// 
			this.chkSubGrupo.Properties.Caption = "Todo SubGrupo";
			this.chkSubGrupo.Size = new System.Drawing.Size(120, 19);
			this.chkSubGrupo.TabIndex = 86;
			this.chkSubGrupo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkSubGrupo_MouseUp);
			this.chkSubGrupo.CheckedChanged += new System.EventHandler(this.chkSubGrupo_CheckedChanged);
			// 
			// chkClase
			// 
			this.chkClase.EditValue = true;
			this.chkClase.Location = new System.Drawing.Point(216, 392);
			this.chkClase.Name = "chkClase";
			// 
			// chkClase.Properties
			// 
			this.chkClase.Properties.Caption = "Toda Clase";
			this.chkClase.Size = new System.Drawing.Size(120, 19);
			this.chkClase.TabIndex = 92;
			this.chkClase.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkClase_MouseUp);
			this.chkClase.CheckedChanged += new System.EventHandler(this.chkClase_CheckedChanged);
			// 
			// chkColor
			// 
			this.chkColor.EditValue = true;
			this.chkColor.Location = new System.Drawing.Point(216, 367);
			this.chkColor.Name = "chkColor";
			// 
			// chkColor.Properties
			// 
			this.chkColor.Properties.Caption = "Todo Color";
			this.chkColor.Size = new System.Drawing.Size(120, 19);
			this.chkColor.TabIndex = 91;
			this.chkColor.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkColor_MouseUp);
			this.chkColor.CheckedChanged += new System.EventHandler(this.chkColor_CheckedChanged);
			// 
			// chkMarca
			// 
			this.chkMarca.EditValue = true;
			this.chkMarca.Location = new System.Drawing.Point(216, 417);
			this.chkMarca.Name = "chkMarca";
			// 
			// chkMarca.Properties
			// 
			this.chkMarca.Properties.Caption = "Toda Marca";
			this.chkMarca.Size = new System.Drawing.Size(120, 19);
			this.chkMarca.TabIndex = 95;
			this.chkMarca.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkMarca_MouseUp);
			this.chkMarca.CheckedChanged += new System.EventHandler(this.chkMarca_CheckedChanged);
			// 
			// cdvClienteCod
			// 
			this.cdvClienteCod.BindingContextCtrl = this;
			this.cdvClienteCod.DataSet = this.cdsCliente;
			this.cdvClienteCod.Sort = "Codigo";
			this.cdvClienteCod.TableName = "";
			this.cdvClienteCod.TableViewName = "Cliente";
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
			// chkSRI
			// 
			this.chkSRI.EditValue = true;
			this.chkSRI.Location = new System.Drawing.Point(216, 442);
			this.chkSRI.Name = "chkSRI";
			// 
			// chkSRI.Properties
			// 
			this.chkSRI.Properties.Caption = "Toda Grupo SRI";
			this.chkSRI.Size = new System.Drawing.Size(120, 19);
			this.chkSRI.TabIndex = 97;
			this.chkSRI.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkSRI_MouseUp);
			this.chkSRI.CheckedChanged += new System.EventHandler(this.chkSRI_CheckedChanged);
			// 
			// cmbImpuesto
			// 
			this.cmbImpuesto.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todo Impuesto";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Impuesto Mayor a Cero";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Impuesto Cero";
			this.cmbImpuesto.Items.Add(valueListItem1);
			this.cmbImpuesto.Items.Add(valueListItem2);
			this.cmbImpuesto.Items.Add(valueListItem3);
			this.cmbImpuesto.Location = new System.Drawing.Point(24, 121);
			this.cmbImpuesto.Name = "cmbImpuesto";
			this.cmbImpuesto.Size = new System.Drawing.Size(160, 22);
			this.cmbImpuesto.TabIndex = 99;
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
			// 
			// chkCredTribut
			// 
			this.chkCredTribut.EditValue = true;
			this.chkCredTribut.Location = new System.Drawing.Point(216, 467);
			this.chkCredTribut.Name = "chkCredTribut";
			// 
			// chkCredTribut.Properties
			// 
			this.chkCredTribut.Properties.Caption = "Toda Créd. Trib.";
			this.chkCredTribut.Size = new System.Drawing.Size(120, 19);
			this.chkCredTribut.TabIndex = 100;
			this.chkCredTribut.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkCredTribut_MouseUp);
			this.chkCredTribut.CheckedChanged += new System.EventHandler(this.chkCredTribut_CheckedChanged);
			// 
			// udsBodegas
			// 
			ultraDataColumn2.DataType = typeof(int);
			this.udsBodegas.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn1,
																													 ultraDataColumn2});
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoRepFact.Bodega"));
			this.cmbBodega.DataSource = this.udsBodegas;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(120, 95);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(88, 21);
			this.cmbBodega.TabIndex = 128;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// chkNeto
			// 
			this.chkNeto.Location = new System.Drawing.Point(128, 17);
			this.chkNeto.Name = "chkNeto";
			// 
			// chkNeto.Properties
			// 
			this.chkNeto.Properties.Caption = "Neto";
			this.chkNeto.Size = new System.Drawing.Size(75, 19);
			this.chkNeto.TabIndex = 130;
			// 
			// chkTipo
			// 
			this.chkTipo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoRepFact.bTipo"));
			this.chkTipo.EditValue = true;
			this.chkTipo.Location = new System.Drawing.Point(216, 292);
			this.chkTipo.Name = "chkTipo";
			// 
			// chkTipo.Properties
			// 
			this.chkTipo.Properties.Caption = "Todo Tipo";
			this.chkTipo.Size = new System.Drawing.Size(120, 19);
			this.chkTipo.TabIndex = 131;
			this.chkTipo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkTipo_MouseUp);
			this.chkTipo.CheckedChanged += new System.EventHandler(this.chkTipo_CheckedChanged);
			// 
			// cmbSubGrupo
			// 
			this.cmbSubGrupo.Location = new System.Drawing.Point(384, 317);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			// 
			// cmbSubGrupo.Properties
			// 
			this.cmbSubGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSubGrupo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("SubGrupo", "SubGrupo", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSubGrupo.Properties.DisplayMember = "SubGrupo";
			this.cmbSubGrupo.Properties.NullText = "";
			this.cmbSubGrupo.Properties.PopupWidth = 400;
			this.cmbSubGrupo.Properties.ValueMember = "idSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(176, 20);
			this.cmbSubGrupo.TabIndex = 133;
			// 
			// cmbModelo
			// 
			this.cmbModelo.Location = new System.Drawing.Point(384, 342);
			this.cmbModelo.Name = "cmbModelo";
			// 
			// cmbModelo.Properties
			// 
			this.cmbModelo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbModelo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Modelo", "Modelo", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbModelo.Properties.DisplayMember = "Modelo";
			this.cmbModelo.Properties.NullText = "";
			this.cmbModelo.Properties.PopupWidth = 400;
			this.cmbModelo.Properties.ValueMember = "idModelo";
			this.cmbModelo.Size = new System.Drawing.Size(176, 20);
			this.cmbModelo.TabIndex = 134;
			// 
			// cmbColor
			// 
			this.cmbColor.Location = new System.Drawing.Point(384, 367);
			this.cmbColor.Name = "cmbColor";
			// 
			// cmbColor.Properties
			// 
			this.cmbColor.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbColor.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Color", "Color", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbColor.Properties.DisplayMember = "Color";
			this.cmbColor.Properties.NullText = "";
			this.cmbColor.Properties.PopupWidth = 400;
			this.cmbColor.Properties.ValueMember = "idColor";
			this.cmbColor.Size = new System.Drawing.Size(176, 20);
			this.cmbColor.TabIndex = 135;
			// 
			// cmbClase
			// 
			this.cmbClase.Location = new System.Drawing.Point(384, 392);
			this.cmbClase.Name = "cmbClase";
			// 
			// cmbClase.Properties
			// 
			this.cmbClase.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbClase.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Clase", "Clase", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbClase.Properties.DisplayMember = "Clase";
			this.cmbClase.Properties.NullText = "";
			this.cmbClase.Properties.PopupWidth = 400;
			this.cmbClase.Properties.ValueMember = "idClase";
			this.cmbClase.Size = new System.Drawing.Size(176, 20);
			this.cmbClase.TabIndex = 136;
			// 
			// cmbCredTribut
			// 
			this.cmbCredTribut.Location = new System.Drawing.Point(384, 467);
			this.cmbCredTribut.Name = "cmbCredTribut";
			// 
			// cmbCredTribut.Properties
			// 
			this.cmbCredTribut.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCredTribut.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CredTributario", "CredTributario", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCredTribut.Properties.DisplayMember = "CredTributario";
			this.cmbCredTribut.Properties.NullText = "";
			this.cmbCredTribut.Properties.PopupWidth = 400;
			this.cmbCredTribut.Properties.ValueMember = "idCredTributario";
			this.cmbCredTribut.Size = new System.Drawing.Size(176, 20);
			this.cmbCredTribut.TabIndex = 139;
			// 
			// cmbSRI
			// 
			this.cmbSRI.Location = new System.Drawing.Point(384, 442);
			this.cmbSRI.Name = "cmbSRI";
			// 
			// cmbSRI.Properties
			// 
			this.cmbSRI.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																	 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSRI.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																			 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Detalle", "Detalle", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSRI.Properties.DisplayMember = "Detalle";
			this.cmbSRI.Properties.NullText = "";
			this.cmbSRI.Properties.PopupWidth = 400;
			this.cmbSRI.Properties.ValueMember = "idArticuloSRI";
			this.cmbSRI.Size = new System.Drawing.Size(176, 20);
			this.cmbSRI.TabIndex = 138;
			// 
			// cmbMarca
			// 
			this.cmbMarca.Location = new System.Drawing.Point(384, 417);
			this.cmbMarca.Name = "cmbMarca";
			// 
			// cmbMarca.Properties
			// 
			this.cmbMarca.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbMarca.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Marca", "Marca", 31, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbMarca.Properties.DisplayMember = "Marca";
			this.cmbMarca.Properties.NullText = "";
			this.cmbMarca.Properties.PopupWidth = 400;
			this.cmbMarca.Properties.ValueMember = "idMarca";
			this.cmbMarca.Size = new System.Drawing.Size(176, 20);
			this.cmbMarca.TabIndex = 137;
			// 
			// RepArticulos
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(584, 510);
			this.Controls.Add(this.cmbCredTribut);
			this.Controls.Add(this.cmbSRI);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.cmbClase);
			this.Controls.Add(this.cmbColor);
			this.Controls.Add(this.cmbModelo);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.chkTipo);
			this.Controls.Add(this.chkNeto);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.chkCredTribut);
			this.Controls.Add(this.cmbImpuesto);
			this.Controls.Add(this.chkSRI);
			this.Controls.Add(this.chkMarca);
			this.Controls.Add(this.chkClase);
			this.Controls.Add(this.chkColor);
			this.Controls.Add(this.chkModelo);
			this.Controls.Add(this.chkSubGrupo);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.cmbArtCod);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbGrupoArt);
			this.Controls.Add(this.chkGrupoArt);
			this.Controls.Add(this.chkArticulo);
			this.Controls.Add(this.chkBodega);
			this.Controls.Add(this.spnBodega);
			this.Controls.Add(this.cmbSector);
			this.Controls.Add(this.chkSector);
			this.Controls.Add(this.cmbCiudad);
			this.Controls.Add(this.chkCiudad);
			this.Controls.Add(this.chkAnulado);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.chkTipoTrans);
			this.Controls.Add(this.cmbComprobante);
			this.Controls.Add(this.chkComprobante);
			this.Controls.Add(this.rdgFecha);
			this.Controls.Add(this.cmbSubCentro);
			this.Controls.Add(this.chkSubCentro);
			this.Controls.Add(this.cmbCentro);
			this.Controls.Add(this.chkCentro);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.chkUsuario);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.chkGrupo);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.cmbFechas);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbTipoFact);
			this.MaximizeBox = false;
			this.Name = "RepArticulos";
			this.Text = "Reportes de Artículos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepArticulos_Closing);
			this.Load += new System.EventHandler(this.RepFacturacion_Load);
			this.Activated += new System.EventHandler(this.RepFacturacion_Activated);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFact.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkUsuario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubCentro.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipoTrans.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkResumen.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkAnulado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCiudad.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSector.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkBodega.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoArt.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkGrupoArt.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkArticulo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArtCod.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkModelo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkClase.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkColor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkMarca.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvClienteCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSRI.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbImpuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCredTribut.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNeto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTipo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCredTribut.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSRI.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca.Properties)).EndInit();
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
			this.errorProvider.SetError(this.cmbSector, "");
			this.errorProvider.SetError(this.cmbSubCentro, "");
			this.errorProvider.SetError(this.cmbUsuario, "");
			this.errorProvider.SetError(this.cmbArtCod, "");
			this.errorProvider.SetError(this.cmbGrupoArt, "");
			this.errorProvider.SetError(this.cmbSubGrupo, "");
			this.errorProvider.SetError(this.cmbModelo, "");
			this.errorProvider.SetError(this.cmbColor, "");
			this.errorProvider.SetError(this.cmbClase, "");
			this.errorProvider.SetError(this.cmbMarca, "");
			this.errorProvider.SetError(this.cmbSRI, "");
		
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
			if (chkUsuario.Checked == false || chkUsuario.CheckState == CheckState.Indeterminate)
				if (cmbUsuario.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbUsuario, "Ingrese Vendedor");
					bOk = false;
				}
			if (chkCentro.Checked == false || chkCentro.CheckState == CheckState.Indeterminate)
				if (cmbCentro.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCentro, "Ingrese Centro de Costo");
					bOk = false;
				}
			if (chkSubCentro.Checked == false || chkSubCentro.CheckState == CheckState.Indeterminate)
				if (cmbSubCentro.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSubCentro, "Ingrese Proyecto");
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
			if (chkSector.Checked == false || chkSector.CheckState == CheckState.Indeterminate)
				if (cmbSector.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSector, "Ingrese Sector");
					bOk = false;
				}
			if (chkArticulo.Checked == false || chkArticulo.CheckState == CheckState.Indeterminate)
				if (cmbArtCod.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbArtCod, "Ingrese Artículo");
					bOk = false;
				}
			if (chkGrupoArt.Checked == false || chkGrupoArt.CheckState == CheckState.Indeterminate)
				if (cmbGrupoArt.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbGrupoArt, "Ingrese Grupo");
					bOk = false;
				}
			if (chkBodega.Checked == false || chkBodega.CheckState == CheckState.Indeterminate)
			{
				#region Numero de Bodega
				int iBodega = 0;
				if (drSeteoF["VerBodegaNombre"] != DBNull.Value && (bool) drSeteoF["VerBodegaNombre"] == true)
				{
					iBodega = int.Parse(this.spnBodega.EditValue.ToString());
				}
				else
				{
					if (spnBodega.Value > 0)
						iBodega = (int) cmbBodega.Value;
				}
				#endregion Numero de Bodega

				if (iBodega < 1)
				{
					errorProvider.SetError(spnBodega, "Ingrese Bodega");
					bOk = false;
				}
			}
			if (chkSubGrupo.Checked == false || chkSubGrupo.CheckState == CheckState.Indeterminate)
				if (cmbSubGrupo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSubGrupo, "Ingrese SubGrupo");
					bOk = false;
				}
			if (chkModelo.Checked == false || chkModelo.CheckState == CheckState.Indeterminate)
				if (cmbModelo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbModelo, "Ingrese Modelo");
					bOk = false;
				}
			if (chkColor.Checked == false || chkColor.CheckState == CheckState.Indeterminate)
				if (cmbColor.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbColor, "Ingrese Color");
					bOk = false;
				}
			if (chkClase.Checked == false || chkClase.CheckState == CheckState.Indeterminate)
				if (cmbClase.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbClase, "Ingrese Clase");
					bOk = false;
				}
			if (chkMarca.Checked == false || chkMarca.CheckState == CheckState.Indeterminate)
				if (cmbMarca.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbMarca, "Ingrese Marca");
					bOk = false;
				}
			if (chkSRI.Checked == false || chkSRI.CheckState == CheckState.Indeterminate)
				if (cmbSRI.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSRI, "Ingrese Grupo del SRI");
					bOk = false;
				}
			if (chkCredTribut.Checked == false || chkCredTribut.CheckState == CheckState.Indeterminate)
				if (cmbCredTribut.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCredTribut, "Ingrese Credito Tribuario");
					bOk = false;
				}
			return bOk;
		}

		CultureInfo us = new CultureInfo("en-US");
		private string[] Filtro()
		{
			DateTime dtFechaIni = (DateTime) this.dtFechaDesde.EditValue;
			DateTime dtFechaFin = (DateTime) this.dtFechaHasta.EditValue;

			string stFiltro = "";
			string stTitulo = "";
			string stFecha = "Fecha";
			if ((int) this.rdgFecha.EditValue == 1) stFecha = "FechaVencimiento";
			if ((int) this.rdgFecha.EditValue == 2) stFecha = "FechaEntrega";

			stFiltro = "{Compra." + stFecha + "} >= #"
				+ dtFechaIni.ToString("MM/dd/yyyy", us) 
				+ "# And {Compra." + stFecha + "} < #" 
				+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy", us) + "#";
			stTitulo = "Desde " + dtFechaIni.ToString("dd/MMM/yyyy") +
				" hasta "+ dtFechaFin.ToString("dd/MMM/yyyy");

			string stNoFiltroRep = cmbReporte.ActiveRow.Cells["NoFiltro"].Value.ToString();
			int iLong = stNoFiltroRep.IndexOf("NF");
			if (iLong >= 0) // no use Fechas
			{
				stFiltro = "true";
				stTitulo = "";
			}	

			// No usa los filtros si no se veririca Compras
			iLong = stNoFiltroRep.IndexOf("NC"); // No compra
			if (iLong < 0)
			{
				if (chkTipoTrans.Checked == false)
				{
					int iTipoF = int.Parse(this.cmbTipoFact.EditValue.ToString());
					if (chkNeto.Checked && (iTipoF == 1 || iTipoF == 4)) 
					{
						if (iTipoF == 1)
							stFiltro += " And {Compra.idTipoFactura} In [1, 5]";
						else
							stFiltro += " And {Compra.idTipoFactura} In [4, 6]";

						stTitulo += "Tipo " + cmbTipoFact.Text.Trim() + " y Dev. ";
					}
					else
					{
						stFiltro += " And {Compra.idTipoFactura} = " + cmbTipoFact.EditValue.ToString();
						stTitulo += "Transaccion " + cmbTipoFact.Text + " ";
					}
				}

				#region Visto en Cliente
				if (chkCliente.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					stFiltro += " And {Compra.idCliente} != " + cmbCliente.EditValue.ToString();
					stTitulo += "Nombre No es " + cmbCliente.Text + " ";
				}
				if (this.chkCliente.Checked == false)
				{
					stFiltro += " And {Compra.idCliente} = " + cmbCliente.EditValue.ToString();
					stTitulo += "Nombre " + cmbCliente.Text + " ";
				}
				#endregion Todo Cliente
				#region Grupo
				if (chkGrupo.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					stFiltro += " And {Cliente.idGrupoCliente} != " + cmbGrupo.EditValue.ToString();
					stTitulo += "Grupo No es " + cmbGrupo.Text + " ";
				}
				if (chkGrupo.Checked == false)
				{
					stFiltro += " And {Cliente.idGrupoCliente} = " + cmbGrupo.EditValue.ToString();
					stTitulo += "Grupo " + cmbGrupo.Text + " ";
				}
				#endregion Grupo
				#region Usuario
				if (chkUsuario.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVB'")) 	
					{
						stFiltro += " And {DetCompra.idVendedor} != " + cmbUsuario.EditValue.ToString();
						stTitulo += "Vendedor No es " + cmbUsuario.Text + " ";
					}
					else
					{
						stFiltro += " And {Compra.idVendedor} != " + cmbUsuario.EditValue.ToString();
						stTitulo += "Vendedor No es " + cmbUsuario.Text + " ";
					}
				}
				if (chkUsuario.Checked == false)
				{
					if (Funcion.bEjecutaSQL(cdsArticulo, "Exec SeteoGExiste 'FCVB'")) 	
					{
						stFiltro += " And {DetCompra.idVendedor} = " + cmbUsuario.EditValue.ToString();
						stTitulo += "Vendedor " + cmbUsuario.Text + " ";
					}
					else
					{
						stFiltro += " And {Compra.idVendedor} = " + cmbUsuario.EditValue.ToString();
						stTitulo += "Vendedor " + cmbUsuario.Text + " ";
					}
				}
				#endregion Usuario
				#region Credito Tributario
				if (chkCredTribut.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					stFiltro += " And {Compra.idCredTributario} != " + cmbCredTribut.EditValue.ToString();
					stTitulo += "Cred Tribut No es " + cmbCredTribut.Text + " ";
				}
				if (chkCredTribut.Checked == false)
				{
					stFiltro += " And {Compra.idCredTributario} = " + cmbCredTribut.EditValue.ToString();
					stTitulo += "Cred Tribut " + cmbCredTribut.Text + " ";
				}
				#endregion
				#region Centro de Costo
				if (chkCentro.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					if (drSeteoF["VerProyecto"] == DBNull.Value || (bool) drSeteoF["VerProyecto"] == true)
						stFiltro += " And {DetCompra.idProyecto} != ";
					else
						stFiltro += " And {Compra.idProyecto} != ";

					stFiltro += cmbCentro.EditValue.ToString();
					stTitulo += "Centro de Costo No es " + cmbCentro.Text + " ";
				}
				if (chkCentro.Checked == false)
				{
					if (drSeteoF["VerProyecto"] == DBNull.Value || (bool) drSeteoF["VerProyecto"] == true)
						stFiltro += " And {DetCompra.idProyecto} = ";
					else
						stFiltro += " And {Compra.idProyecto} = ";

					stFiltro += cmbCentro.EditValue.ToString();
					stTitulo += "Centro de Costo " + cmbCentro.Text + " ";
				}
				#endregion Centro de Costo
				#region Proyecto
				if (chkSubCentro.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					if (drSeteoF["VerSubProyecto"] == DBNull.Value || (bool) drSeteoF["VerSubProyecto"] == true)
						stFiltro += " And {DetCompra.idSubProyecto} != ";
					else
						stFiltro += " And {Compra.idSubProyecto} != ";

					stFiltro += cmbSubCentro.EditValue.ToString();
					stTitulo += "Proyecto No es " + cmbSubCentro.Text + " ";
				}
				if (chkSubCentro.Checked == false)
				{
					if (drSeteoF["VerSubProyecto"] == DBNull.Value || (bool) drSeteoF["VerSubProyecto"] == true)
						stFiltro += " And {DetCompra.idSubProyecto} = ";
					else
						stFiltro += " And {Compra.idSubProyecto} = ";

					stFiltro += cmbSubCentro.EditValue.ToString();
					stTitulo += "Proyecto " + cmbSubCentro.Text + " ";
				}
				#endregion Proyecto
				#region Comprobante
				if (chkComprobante.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					stFiltro += " And {Compra.idComprobante} != " + cmbComprobante.EditValue.ToString();
					stTitulo += "Comprobante No es " + cmbComprobante.Text + " ";
				}
				if (chkComprobante.Checked == false)
				{
					stFiltro += " And {Compra.idComprobante} = " + cmbComprobante.EditValue.ToString();
					stTitulo += "Comprobante " + cmbComprobante.Text + " ";
				}
				#endregion Comprobante
				#region Ciudad
				if (chkCiudad.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					stFiltro += " And {Cliente.idCiudad} != " + cmbCiudad.EditValue.ToString();
					stTitulo += "Ciudad No es " + cmbCiudad.Text + " ";
				}
				if (chkCiudad.Checked == false)
				{
					stFiltro += " And {Cliente.idCiudad} = " + cmbCiudad.EditValue.ToString();
					stTitulo += "Ciudad " + cmbCiudad.Text + " ";
				}
				#endregion Ciudad
				#region Sector
				if (chkSector.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					stFiltro += " And {Cliente.idSector} != " + cmbSector.EditValue.ToString();
					stTitulo += "Sector No es " + cmbSector.Text + " ";
				}
				if (chkSector.Checked == false)
				{
					stFiltro += " And {Cliente.idSector} = " + cmbSector.EditValue.ToString();
					stTitulo += "Sector " + cmbSector.Text + " ";
				}
				#endregion Sector
				#region Bodega
				if (chkBodega.CheckState == System.Windows.Forms.CheckState.Indeterminate)
				{
					if ((bool) drSeteoF["VerBodega"] == false)
					{
						stFiltro += " And {Compra.Bodega} != " + spnBodega.EditValue.ToString();
						stTitulo += "Bodega No es " + spnBodega.Value.ToString() + " ";
					}
					else
					{
						stFiltro += " And {DetCompra.Bodega} = " + spnBodega.EditValue.ToString();
						stTitulo += "Bodega L. No es " + spnBodega.Value.ToString() + " ";
					}
				}
				if (chkBodega.Checked == false)
				{
					if ((bool) drSeteoF["VerBodega"] == false)
					{
						stFiltro += " And {Compra.Bodega} = " + spnBodega.EditValue.ToString();
						stTitulo += "Bodega " + spnBodega.Value.ToString() + " ";
					}
					else
					{
						stFiltro += " And {DetCompra.Bodega} = " + spnBodega.EditValue.ToString();
						stTitulo += "Bodega L. " + spnBodega.Value.ToString() + " ";
					}
				}
				#endregion Bodega
				#region Anulado
				if (this.chkAnulado.Checked)
				{
					stTitulo += " con Anulados ";
				}
				else
				{
					stFiltro += " And {Compra.Borrar} = false" ;
				}
				#endregion Anulado
			} // Fin de No Compra

			#region Articulo
			if (this.chkArticulo.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idArticulo} != " + cmbArtCod.EditValue.ToString();
				stTitulo += "Articulo No es " + cmbArtCod.Text + " ";
			}
			if (this.chkArticulo.Checked == false)
			{
				stFiltro += " And {Articulo.idArticulo} = " + cmbArtCod.EditValue.ToString();
				stTitulo += "Articulo " + cmbArtCod.Text + " ";
			}
			#endregion Articulo
			#region Grupo de Articulo
			if (this.chkGrupoArt.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idGrupoArticulo} != " + cmbGrupoArt.EditValue.ToString();
				stTitulo += "Grupo Articulos No es " + cmbGrupoArt.Text + " ";
			}
			if (this.chkGrupoArt.Checked == false)
			{
				stFiltro += " And {Articulo.idGrupoArticulo} = " + cmbGrupoArt.EditValue.ToString();
				stTitulo += "Grupo Articulos " + cmbGrupoArt.Text + " ";
			}
			#endregion Grupo de Articulo
			#region SubGrupo de Articulo
			if (this.chkSubGrupo.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idSubGrupo} != " + cmbSubGrupo.EditValue.ToString();
				stTitulo += "SubGrupo Articulos No es " + cmbSubGrupo.Text + " ";
			}
			if (this.chkSubGrupo.Checked == false)
			{
				stFiltro += " And {Articulo.idSubGrupo} = " + cmbSubGrupo.EditValue.ToString();
				stTitulo += "SubGrupo Articulos " + cmbSubGrupo.Text + " ";
			}
			#endregion
			#region Tipo
			if (this.chkTipo.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idTipoGrupo} <> 3  And {Articulo.idTipoGrupo} <> 5 ";
				stTitulo += "Tipo No Servicio ni AF ";
			}
			else
			{
				if (this.chkTipo.Checked == false)
				{
					stFiltro += " And {Articulo.idTipoGrupo} = " + cmbTipo.EditValue.ToString();
					stTitulo += "Tipo " + cmbTipo.Text + " ";
				}
			}
			#endregion
			#region Modelo
			if (this.chkModelo.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idModelo} != " + cmbModelo.EditValue.ToString();
				stTitulo += "Modelo No es " + cmbModelo.Text + " ";
			}
			if (this.chkModelo.Checked == false)
			{
				stFiltro += " And {Articulo.idModelo} = " + cmbModelo.EditValue.ToString();
				stTitulo += "Modelo " + cmbModelo.Text + " ";
			}
			#endregion Modelo
			#region Color
			if (this.chkColor.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idColor} != " + cmbColor.EditValue.ToString();
				stTitulo += "Color No es " + cmbColor.Text + " ";
			}
			if (this.chkColor.Checked == false)
			{
				stFiltro += " And {Articulo.idColor} = " + cmbColor.EditValue.ToString();
				stTitulo += "Color " + cmbColor.Text + " ";
			}
			#endregion Color
			#region Clase
			if (this.chkClase.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idClase} != " + cmbClase.EditValue.ToString();
				stTitulo += "Clase No es " + cmbClase.Text + " ";
			}
			if (this.chkClase.Checked == false)
			{
				stFiltro += " And {Articulo.idClase} = " + cmbClase.EditValue.ToString();
				stTitulo += "Clase " + cmbClase.Text + " ";
			}
			#endregion Clase
			#region Marca
			if (this.chkMarca.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idMarca} != " + cmbMarca.EditValue.ToString();
				stTitulo += "Marca No es " + cmbMarca.Text + " ";
			}
			if (this.chkMarca.Checked == false)
			{
				stFiltro += " And {Articulo.idMarca} = " + cmbMarca.EditValue.ToString();
				stTitulo += "Marca " + cmbMarca.Text + " ";
			}
			#endregion Marca
			#region Sri
			if (chkSRI.CheckState == System.Windows.Forms.CheckState.Indeterminate)
			{
				stFiltro += " And {Articulo.idArticuloSRI} != " + cmbSRI.EditValue.ToString();
				stTitulo += "Grupo SRI No es " + cmbSRI.Text + " ";
			}
			if (chkSRI.Checked == false)
			{
				stFiltro += " And {Articulo.idArticuloSRI} = " + cmbSRI.EditValue.ToString();
				stTitulo += "Grupo SRI " + cmbSRI.Text + " ";
			}
			#endregion Sri
			#region Impuesto
			if ((int) cmbImpuesto.Value == 1)
			{
				stFiltro += " And {DetCompra.Impuesto} > 0 " ;
				stTitulo += "Con Impuesto ";
			}
			if ((int) cmbImpuesto.Value == 2)
			{
				stFiltro += " And {DetCompra.Impuesto} = 0 " ;
				stTitulo += "Sin Impuesto ";
			}
			#endregion Impuesto
			
			string[] stFiltro1 = new string[2];
			stFiltro1[0] = stFiltro;
			stFiltro1[1] = stTitulo;

			return stFiltro1;
		}

		private bool Seguridad()
		{
			bool bOk = true;
			errorProvider.SetError(chkTipoTrans, "");
			errorProvider.SetError(cmbTipoFact, "");
			if (chkTipoTrans.Checked) 
			{
				string stSelect = "Select Count(*) From seguridad Where Id_Seg >= 107"
					+ " And Id_Seg <= 121 And n_" + MenuLatinium.iNivel.ToString() + " = 0";
				int iCuenta = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
				if (iCuenta > 0)
				{
					bOk = false;
					errorProvider.SetError(chkTipoTrans, "Seleccione un tipo de transaccion");
				}
			}
			#region Filtro de Facturas
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
							errorProvider.SetError(cmbTipoFact, "No tiene acceso a Cotizaciones");
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
			#endregion
			return bOk;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Seguridad()) return;
			if (!Verifica()) return;
			DateTime dtIni = DateTime.Now;

			Cursor = Cursors.WaitCursor;
			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'GEUR'"))
			{
				try
				{
					this.cdsSeteoF.Update();
				}
				catch
				{
				}
			}
			dr = this.cdsSeteoF.TableViews["SeteoRepFact"].Rows[0];
			string[] stFiltro1 = new string[2];
			stFiltro1 = Filtro();
			string stFiltro = stFiltro1[0];
			string stTitulo = stFiltro1[1];
			string stReporte = cmbReporte.Text.Trim() + ".rpt";
			string stFiltroRep = cmbReporte.ActiveRow.Cells["Filtro"].Value.ToString();
			if (stFiltroRep != null) stFiltro += stFiltroRep;

			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(stTitulo);
			if (this.chkResumen.Checked) miReporte.Resumen();
			miReporte.Show();
			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			btAceptar.ToolTip = "Tiempo de ultimo reporte : " + ts.TotalSeconds.ToString("n2");
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
					dtDesde = DateTime.Today;
					break;
				case 1: // Solo Hoy
					dtDesde = DateTime.Today;
					dtHasta = DateTime.Today;
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
						+ DateTime.Today.Year.ToString(), us);
					break;
				case 5: // El año pasado
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
				case 11: // Periodo Digitado
					break;
			}
			dr["FechaIni"] = dtDesde;
			dr["FechaFin"] = dtHasta;
//			this.cdsSeteoF.Update();
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCliente.Enabled = !this.chkCliente.Checked;
		}

		private void chkGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbGrupo.Enabled = !this.chkGrupo.Checked;		
		}

		private void chkUsuario_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbUsuario.Enabled = !this.chkUsuario.Checked;
		}

		private void chkCentro_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbCentro.Enabled = !this.chkCentro.Checked;
		}

		private void chkSubCentro_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbSubCentro.Enabled = !this.chkSubCentro.Checked;
		}

		private void chkComprobante_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbComprobante.Enabled = !this.chkComprobante.Checked;
		}

		private void chkCiudad_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbCiudad.Enabled = !this.chkCiudad.Checked;
		}

		private void chkSector_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbSector.Enabled = !this.chkSector.Checked;
		}

		private void chkTipoTrans_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTipoFact.Enabled = !this.chkTipoTrans.Checked;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void chkBodega_CheckedChanged(object sender, System.EventArgs e)
		{
			this.spnBodega.Enabled = !this.chkBodega.Checked;
			this.cmbBodega.Enabled = !this.chkBodega.Checked;
		}

		C1.Data.C1DataRow dr;
		C1.Data.C1DataRow drSeteoF;
		private void RepFacturacion_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepArticulos'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			if (Funcion.bEjecutaSQL(cdsCiudad, "Exec SeteoGExiste 'FTMTE'")) 				
			{
				chkArticulo.Properties.AllowGrayed = true;
				chkBodega.Properties.AllowGrayed = true;
				chkCentro.Properties.AllowGrayed = true;
				chkCiudad.Properties.AllowGrayed = true;
				chkClase.Properties.AllowGrayed = true;
				chkCliente.Properties.AllowGrayed = true;
				chkColor.Properties.AllowGrayed = true;
				chkComprobante.Properties.AllowGrayed = true;
				chkCredTribut.Properties.AllowGrayed = true;
				chkGrupo.Properties.AllowGrayed = true;
				chkGrupoArt.Properties.AllowGrayed = true;
				chkMarca.Properties.AllowGrayed = true;
				chkModelo.Properties.AllowGrayed = true;
				chkSector.Properties.AllowGrayed = true;
				chkSRI.Properties.AllowGrayed = true;
				chkSubCentro.Properties.AllowGrayed = true;
				chkSubGrupo.Properties.AllowGrayed = true;
				chkTipo.Properties.AllowGrayed = true;
				chkTipoTrans.Properties.AllowGrayed = true;
				chkUsuario.Properties.AllowGrayed = true;
			}
			Cursor = Cursors.Default;
			dr = this.cdsSeteoF.TableViews["SeteoRepFact"].Rows[0];
			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			if (dr["FiltroFecha"] == System.DBNull.Value)
				cmbFechas.SelectedIndex = 0;
			else
				this.cmbFechas.SelectedIndex = (int) dr["FiltroFecha"];
			cmbTipoFact.EditValue = 1;
			cmbReporte.Value = 4;
			cmbImpuesto.Value = 0;

			#region Bodega por Nombre
			if (drSeteoF["VerBodegaNombre"] != DBNull.Value && (bool) drSeteoF["VerBodegaNombre"] == true)
			{
				spnBodega.Width = 0;
				udsBodegas.Rows.Clear();
				string stExecBodega = "Exec BodegaUsuario 0"	;//+ IdTipoFactura.ToString();
				SqlDataReader drBod = Funcion.rEscalarSQL(cdsCliente, stExecBodega, true);
				int iNumFilas = 0;
				while(drBod.Read())
				{
					udsBodegas.Rows.Add();
					udsBodegas.Rows[iNumFilas]["Nombre"]=drBod.GetString(0);
					udsBodegas.Rows[iNumFilas]["Bodega"]=drBod.GetInt32(1);
					iNumFilas++;
				}
			}
			else
			{
				cmbBodega.Width = 0;
			}
			#endregion Bodega por Nombre
			#region Visibilidad de Combos
			chkArticulo_CheckedChanged(this, e);
			chkBodega_CheckedChanged(this, e);
			chkCentro_CheckedChanged(this, e);
			chkCiudad_CheckedChanged(this, e);
			chkClase_CheckedChanged(this, e);
			chkCliente_CheckedChanged(this, e);
			chkColor_CheckedChanged(this, e);
			chkComprobante_CheckedChanged(this, e);
			chkCredTribut_CheckedChanged(this, e);
			chkGrupo_CheckedChanged(this, e);
			chkGrupoArt_CheckedChanged(this, e);
			chkMarca_CheckedChanged(this, e);
			chkModelo_CheckedChanged(this, e);
			chkSector_CheckedChanged(this, e);
			chkSRI_CheckedChanged(this, e);
			chkSubCentro_CheckedChanged(this, e);
			chkSubGrupo_CheckedChanged(this, e);
			chkTipo_CheckedChanged(this, e);
			chkTipoTrans_CheckedChanged(this, e);
			chkUsuario_CheckedChanged(this, e);
			#endregion Visibilidad de Combos
			#region cargar Combos
			cmbSubGrupo.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idSubGrupo, SubGrupo From ArticuloSubGrupo");
			cmbModelo.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idModelo, Modelo From ArticuloModelo");
			cmbColor.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idColor, Color From ArticuloColor");
			cmbClase.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idClase, Clase From ArticuloClase");
			cmbMarca.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idMarca, Marca From ArticuloMarca");
			cmbSRI.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idArticuloSRI, Detalle From ArticuloSRI");
			cmbCredTribut.Properties.DataSource = Funcion.dvProcedimiento(cdsArticulo,
				"Select idCredTributario, CredTributario From CompraCredTribut");
			#endregion cargar Combos
		}

		private void RepFacturacion_Activated(object sender, System.EventArgs e)
		{
			this.cmbTipoFact.Focus();
		}

		private void chkArticulo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbArtCod.Enabled = !this.chkArticulo.Checked;
		}

		private void chkGrupoArt_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbGrupoArt.Enabled = !this.chkGrupoArt.Checked;
		}

		private void chkTipo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTipo.Enabled = !this.chkTipo.Checked;		
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

		private void cdsTipo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTipo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsGrupoArt_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupoArt.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsReporte_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void chkSubGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbSubGrupo.Enabled = !chkSubGrupo.Checked;
			if (!chkSubGrupo.Checked) cmbSubGrupo.Select();
		}

		private void chkModelo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbModelo.Enabled = !chkModelo.Checked;
			if (!chkModelo.Checked) cmbModelo.Select();		
		}

		private void chkColor_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbColor.Enabled = !chkColor.Checked;
			if (!chkColor.Checked) cmbColor.Select();
		}

		private void chkClase_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbClase.Enabled = !chkClase.Checked;
			if (!chkClase.Checked) cmbClase.Select();
		}

		private void chkMarca_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbMarca.Enabled = !chkMarca.Checked;
			if (!chkMarca.Checked) cmbMarca.Select();
		}

		private void cmbCliente_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (this.cmbCliente.Properties.DisplayMember == "Codigo")
			{
				this.cmbCliente.Properties.Columns.Clear();
				this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default)});
				this.cmbCliente.Properties.DataSource = this.cdvCliente;
				this.cmbCliente.Properties.DisplayMember = "Nombre";
			}
			else
			{
				this.cmbCliente.Properties.Columns.Clear();
				this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default),
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
				this.cmbCliente.Properties.DataSource = this.cdvClienteCod;
				this.cmbCliente.Properties.DisplayMember = "Codigo";
			}
		}

		private void cmbArtCod_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (this.cmbArtCod.Properties.DisplayMember == "Codigo")
			{
				this.cmbArtCod.Properties.Columns.Clear();
				this.cmbArtCod.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Artículo", 180, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 80, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
				this.cmbArtCod.Properties.DataSource = this.cdvArticulo;
				this.cmbArtCod.Properties.DisplayMember = "Articulo";
			}
			else
			{
				this.cmbArtCod.Properties.Columns.Clear();
				this.cmbArtCod.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 80, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Articulo", "Artículo", 180, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
				this.cmbArtCod.Properties.DataSource = this.cdvArtCodigo;
				this.cmbArtCod.Properties.DisplayMember = "Codigo";
			}		
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void chkSRI_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbSRI.Enabled = !chkSRI.Checked;
			if (!chkSRI.Checked) cmbSRI.Select();		
		}

		private void RepArticulos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepArticulos'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void chkCredTribut_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCredTribut.Enabled = !chkCredTribut.Checked;
			if (!chkCredTribut.Checked) cmbCredTribut.Select();
		}

		private void chkTipo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button!= MouseButtons.Right) return;
			chkTipo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbTipo.Enabled = true;
			cmbTipo.SelectAll();
		}

		private void chkCliente_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCliente.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCliente.Enabled = true;
			cmbCliente.Select();
		}

		private void chkGrupo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkGrupo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbGrupo.Enabled = true;
			cmbGrupo.Select();
		}

		private void chkUsuario_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkUsuario.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbUsuario.Enabled = true;
			cmbUsuario.Select();
		}

		private void chkCentro_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCentro.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCentro.Enabled = true;
			cmbCentro.Select();
		}

		private void chkSubCentro_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkSubCentro.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbSubCentro.Enabled = true;
			cmbSubCentro.Select();
		}

		private void chkComprobante_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkComprobante.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbComprobante.Enabled = true;
			cmbComprobante.Select();
		}

		private void chkCiudad_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCiudad.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCiudad.Enabled = true;
			cmbCiudad.Select();
		}

		private void chkSector_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkSector.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbSector.Enabled = true;
			cmbSector.Select();
		}

		private void chkArticulo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkArticulo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbArtCod.Enabled = true;
			cmbArtCod.Select();
		}

		private void chkGrupoArt_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkGrupoArt.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbGrupoArt.Enabled = true;
			cmbGrupoArt.Select();
		}

		private void chkSubGrupo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkSubGrupo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbSubGrupo.Enabled = true;
			cmbSubGrupo.Select();
		}

		private void chkModelo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkModelo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbModelo.Enabled = true;
			cmbModelo.Select();
		}

		private void chkColor_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkColor.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbColor.Enabled = true;
			cmbColor.Select();
		}

		private void chkClase_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkClase.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbClase.Enabled = true;
			cmbClase.Select();
		}

		private void chkMarca_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkMarca.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbMarca.Enabled = true;
			cmbMarca.Select();
		}

		private void chkSRI_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkSRI.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbSRI.Enabled = true;
			cmbSRI.Select();
		}

		private void chkCredTribut_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkCredTribut.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbCredTribut.Enabled = true;
			cmbCredTribut.Select();
		}
	}
}
