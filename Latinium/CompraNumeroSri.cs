using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraNumeroSri.
	/// </summary>
	public class CompraNumeroSri : DevExpress.XtraEditors.XtraForm
	{
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataView cdvSubProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet cdsNumeracion;
		private C1.Data.C1DataSet cdsTipofactura;
		private C1.Data.C1DataView cdvTipoFactura;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.Label lblAutorizacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtActivacionAutoImpresor;
		private Infragistics.Win.Misc.UltraButton btAutoImpresor;
		private System.Windows.Forms.Label label9;
		private C1.Data.C1DataSet cdsAutoImpresor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoTramite;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutorizaSri;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmdTipoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDoc;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProy;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentroC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdFacturaSri;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chv2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chv1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chv5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chv4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chv3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnNumero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnPtoEstablec;
		private System.Windows.Forms.Label label2;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodEstab;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnPtoEmision;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chAprobado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraNumeroSri()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SRITramite", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSRITramite");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tramite");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("FacturaNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFacturaNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura", -1, "cmdTipoF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto", -1, "cmbCentroC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubProyecto", -1, "cmbProy");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutoImpresor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTramite");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDocumento", -1, "cmbDoc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AutorizaSri");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Establecimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntoEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFacturaSri");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SRITipoDoc", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSRITipoDoc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(CompraNumeroSri));
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsNumeracion = new C1.Data.C1DataSet();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsTipofactura = new C1.Data.C1DataSet();
			this.cdvTipoFactura = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.lblAutorizacion = new System.Windows.Forms.Label();
			this.txtActivacionAutoImpresor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAutoImpresor = new Infragistics.Win.Misc.UltraButton();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbTipoTramite = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsAutoImpresor = new C1.Data.C1DataSet();
			this.txtAutorizaSri = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdFacturaSri = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmdTipoF = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbDoc = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCentroC = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbProy = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chv2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chv1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chv5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chv4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chv3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.spnNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnPtoEstablec = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.txtCodEstab = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.spnPtoEmision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chAprobado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumeracion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipofactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtActivacionAutoImpresor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoTramite)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAutoImpresor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizaSri)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFacturaSri)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdTipoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProy)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPtoEstablec)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodEstab)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPtoEmision)).BeginInit();
			this.SuspendLayout();
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "FacturaSri";
			this.barraDato1.DataNombreId = "idFacturaSri";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsNumeracion;
			this.barraDato1.DataTabla = "FacturaSri";
			this.barraDato1.DataTablaHija = "FacturaNumero";
			this.barraDato1.Location = new System.Drawing.Point(32, 384);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 78;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsNumeracion
			// 
			this.cdsNumeracion.BindingContextCtrl = this;
			this.cdsNumeracion.DataLibrary = "LibFacturacion";
			this.cdsNumeracion.DataLibraryUrl = "";
			this.cdsNumeracion.DataSetDef = "dsNumeracion";
			this.cdsNumeracion.EnforceConstraints = false;
			this.cdsNumeracion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsNumeracion.Name = "cdsNumeracion";
			this.cdsNumeracion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsNumeracion.SchemaDef = null;
			this.cdsNumeracion.BeforeFill += new C1.Data.FillEventHandler(this.cdsNumeracion_BeforeFill);
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.RowFilter = "Visible = true";
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
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
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible = true";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cdsTipofactura
			// 
			this.cdsTipofactura.BindingContextCtrl = this;
			this.cdsTipofactura.DataLibrary = "LibFacturacion";
			this.cdsTipofactura.DataLibraryUrl = "";
			this.cdsTipofactura.DataSetDef = "dsCompraTabla";
			this.cdsTipofactura.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipofactura.Name = "cdsTipofactura";
			this.cdsTipofactura.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipofactura.SchemaDef = null;
			// 
			// cdvTipoFactura
			// 
			this.cdvTipoFactura.BindingContextCtrl = this;
			this.cdvTipoFactura.DataSet = this.cdsTipofactura;
			this.cdvTipoFactura.TableName = "";
			this.cdvTipoFactura.TableViewName = "CompraNumero";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// lblAutorizacion
			// 
			this.lblAutorizacion.Location = new System.Drawing.Point(32, 24);
			this.lblAutorizacion.Name = "lblAutorizacion";
			this.lblAutorizacion.Size = new System.Drawing.Size(96, 16);
			this.lblAutorizacion.TabIndex = 91;
			this.lblAutorizacion.Text = "Autorizacion SRI";
			// 
			// txtActivacionAutoImpresor
			// 
			this.txtActivacionAutoImpresor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtActivacionAutoImpresor.Location = new System.Drawing.Point(632, -8);
			this.txtActivacionAutoImpresor.Name = "txtActivacionAutoImpresor";
			this.txtActivacionAutoImpresor.PasswordChar = '*';
			this.txtActivacionAutoImpresor.Size = new System.Drawing.Size(24, 21);
			this.txtActivacionAutoImpresor.TabIndex = 102;
			// 
			// btAutoImpresor
			// 
			this.btAutoImpresor.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAutoImpresor.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btAutoImpresor.Location = new System.Drawing.Point(688, 0);
			this.btAutoImpresor.Name = "btAutoImpresor";
			this.btAutoImpresor.Size = new System.Drawing.Size(64, 16);
			this.btAutoImpresor.TabIndex = 100;
			this.btAutoImpresor.Text = "Activa AutoImpresor";
			this.btAutoImpresor.Click += new System.EventHandler(this.btAutoImpresor_Click);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(32, 96);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(83, 16);
			this.label9.TabIndex = 104;
			this.label9.Text = "Tipo de Tramite";
			// 
			// cmbTipoTramite
			// 
			this.cmbTipoTramite.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoTramite.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoTramite.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsNumeracion, "FacturaSri.idTramite"));
			this.cmbTipoTramite.DataMember = "SRITramite";
			this.cmbTipoTramite.DataSource = this.cdsAutoImpresor;
			ultraGridBand1.AddButtonCaption = "Proyecto";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 50;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 288;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbTipoTramite.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoTramite.DisplayMember = "Tramite";
			this.cmbTipoTramite.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoTramite.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoTramite.Enabled = false;
			this.cmbTipoTramite.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoTramite.Location = new System.Drawing.Point(160, 88);
			this.cmbTipoTramite.Name = "cmbTipoTramite";
			this.cmbTipoTramite.Size = new System.Drawing.Size(120, 21);
			this.cmbTipoTramite.TabIndex = 106;
			this.cmbTipoTramite.ValueMember = "idSRITramite";
			this.cmbTipoTramite.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoTramite_InitializeLayout);
			// 
			// cdsAutoImpresor
			// 
			this.cdsAutoImpresor.BindingContextCtrl = this;
			this.cdsAutoImpresor.DataLibrary = "LibContabilidad";
			this.cdsAutoImpresor.DataLibraryUrl = "";
			this.cdsAutoImpresor.DataSetDef = "dsAutoImpresor";
			this.cdsAutoImpresor.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAutoImpresor.Name = "cdsAutoImpresor";
			this.cdsAutoImpresor.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAutoImpresor.SchemaDef = null;
			// 
			// txtAutorizaSri
			// 
			this.txtAutorizaSri.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsNumeracion, "FacturaSri.Autorizacion"));
			this.txtAutorizaSri.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutorizaSri.Enabled = false;
			this.txtAutorizaSri.Location = new System.Drawing.Point(160, 24);
			this.txtAutorizaSri.Name = "txtAutorizaSri";
			this.txtAutorizaSri.Size = new System.Drawing.Size(120, 21);
			this.txtAutorizaSri.TabIndex = 110;
			// 
			// txtIdFacturaSri
			// 
			this.txtIdFacturaSri.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsNumeracion, "FacturaSri.idFacturaSri"));
			this.txtIdFacturaSri.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdFacturaSri.Location = new System.Drawing.Point(664, -8);
			this.txtIdFacturaSri.Name = "txtIdFacturaSri";
			this.txtIdFacturaSri.Size = new System.Drawing.Size(24, 21);
			this.txtIdFacturaSri.TabIndex = 111;
			this.txtIdFacturaSri.ValueChanged += new System.EventHandler(this.txtIdFacturaSri_ValueChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "FacturaNumero";
			this.ultraGrid1.DataSource = this.cdsNumeracion;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn4.Header.VisiblePosition = 11;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 12;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 13;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.Caption = "TipoComprobante";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn8.Header.Caption = "CentroCosto";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn9.Header.Caption = "Proyecto";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn13.Header.VisiblePosition = 15;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.Caption = "Documento";
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn15.Width = 83;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 106;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Width = 91;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Width = 90;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 19;
			ultraGridColumn23.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
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
																										 ultraGridColumn23});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
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
			this.ultraGrid1.Location = new System.Drawing.Point(24, 168);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(704, 200);
			this.ultraGrid1.TabIndex = 118;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cmdTipoF
			// 
			this.cmdTipoF.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdTipoF.DataMember = "CompraNumero";
			this.cmdTipoF.DataSource = this.cdsTipofactura;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Width = 173;
			ultraGridColumn26.Header.VisiblePosition = 3;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 4;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 5;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 7;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 8;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 9;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 10;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 11;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 12;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 13;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 14;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 15;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 16;
			ultraGridColumn40.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn40});
			this.cmdTipoF.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmdTipoF.DisplayMember = "Nombre";
			this.cmdTipoF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdTipoF.Location = new System.Drawing.Point(64, 208);
			this.cmdTipoF.Name = "cmdTipoF";
			this.cmdTipoF.Size = new System.Drawing.Size(192, 80);
			this.cmdTipoF.TabIndex = 119;
			this.cmdTipoF.ValueMember = "idTipoFactura";
			this.cmdTipoF.Visible = false;
			// 
			// cmbDoc
			// 
			this.cmbDoc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDoc.DataSource = this.cdsAutoImpresor;
			ultraGridBand4.AddButtonCaption = "CompraComprobante";
			ultraGridColumn41.Header.VisiblePosition = 2;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn42.Width = 75;
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43});
			this.cmbDoc.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbDoc.DisplayMember = "Documento";
			this.cmbDoc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDoc.Location = new System.Drawing.Point(256, 208);
			this.cmbDoc.Name = "cmbDoc";
			this.cmbDoc.Size = new System.Drawing.Size(216, 80);
			this.cmbDoc.TabIndex = 120;
			this.cmbDoc.ValueMember = "idSRITipoDoc";
			this.cmbDoc.Visible = false;
			// 
			// cmbCentroC
			// 
			this.cmbCentroC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroC.DataSource = this.cdvProyecto;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Width = 112;
			ultraGridColumn47.Header.VisiblePosition = 3;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 4;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 5;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 6;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 7;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 8;
			ultraGridColumn52.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52});
			this.cmbCentroC.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCentroC.DisplayMember = "";
			this.cmbCentroC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroC.Location = new System.Drawing.Point(424, 208);
			this.cmbCentroC.Name = "cmbCentroC";
			this.cmbCentroC.Size = new System.Drawing.Size(152, 80);
			this.cmbCentroC.TabIndex = 121;
			this.cmbCentroC.ValueMember = "";
			this.cmbCentroC.Visible = false;
			// 
			// cmbProy
			// 
			this.cmbProy.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProy.DataSource = this.cdvSubProyecto;
			ultraGridColumn53.Header.VisiblePosition = 1;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 2;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn56.Header.VisiblePosition = 3;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 4;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 5;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 6;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 7;
			ultraGridColumn60.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60});
			this.cmbProy.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbProy.DisplayMember = "";
			this.cmbProy.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProy.Location = new System.Drawing.Point(584, 208);
			this.cmbProy.Name = "cmbProy";
			this.cmbProy.Size = new System.Drawing.Size(120, 80);
			this.cmbProy.TabIndex = 122;
			this.cmbProy.ValueMember = "idSubProyecto";
			this.cmbProy.Visible = false;
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			this.btAuditoria.Appearance = appearance6;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(344, 383);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 24);
			this.btAuditoria.TabIndex = 123;
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Enabled = false;
			this.cmbFecha.Format = "MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(160, 56);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(120, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 124;
			this.cmbFecha.Value = new System.DateTime(2009, 1, 17, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(32, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(112, 16);
			this.label1.TabIndex = 125;
			this.label1.Text = "Fecha de Caducidad";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chv2);
			this.groupBox1.Controls.Add(this.chv1);
			this.groupBox1.Controls.Add(this.chv5);
			this.groupBox1.Controls.Add(this.chv4);
			this.groupBox1.Controls.Add(this.chv3);
			this.groupBox1.Location = new System.Drawing.Point(312, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(152, 144);
			this.groupBox1.TabIndex = 128;
			this.groupBox1.TabStop = false;
			// 
			// chv2
			// 
			this.chv2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsNumeracion, "FacturaSri.bRetencion"));
			this.chv2.Location = new System.Drawing.Point(24, 40);
			this.chv2.Name = "chv2";
			this.chv2.TabIndex = 130;
			this.chv2.Text = "Retencion";
			this.chv2.Visible = false;
			// 
			// chv1
			// 
			this.chv1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsNumeracion, "FacturaSri.bVenta"));
			this.chv1.Location = new System.Drawing.Point(24, 16);
			this.chv1.Name = "chv1";
			this.chv1.TabIndex = 129;
			this.chv1.Text = "Venta";
			// 
			// chv5
			// 
			this.chv5.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsNumeracion, "FacturaSri.bGuiaRemision"));
			this.chv5.Location = new System.Drawing.Point(24, 112);
			this.chv5.Name = "chv5";
			this.chv5.TabIndex = 128;
			this.chv5.Text = "Guia de Remision";
			// 
			// chv4
			// 
			this.chv4.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsNumeracion, "FacturaSri.bNotaDebito"));
			this.chv4.Location = new System.Drawing.Point(24, 88);
			this.chv4.Name = "chv4";
			this.chv4.TabIndex = 127;
			this.chv4.Text = "Nota de Debito";
			// 
			// chv3
			// 
			this.chv3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsNumeracion, "FacturaSri.bNotaCredito"));
			this.chv3.Location = new System.Drawing.Point(24, 64);
			this.chv3.Name = "chv3";
			this.chv3.TabIndex = 126;
			this.chv3.Text = "Nota de Credito";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(504, 64);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(112, 16);
			this.label4.TabIndex = 142;
			this.label4.Text = "Puntos de Emision";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(504, 96);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(112, 24);
			this.label3.TabIndex = 141;
			this.label3.Text = "Bloque de Facturas";
			// 
			// spnNumero
			// 
			this.spnNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsNumeracion, "FacturaSri.BloqueFactura"));
			this.spnNumero.Enabled = false;
			this.spnNumero.Location = new System.Drawing.Point(624, 96);
			this.spnNumero.MaskInput = "nnnnnnnnn";
			this.spnNumero.MaxValue = 999999999;
			this.spnNumero.MinValue = 1;
			this.spnNumero.Name = "spnNumero";
			this.spnNumero.PromptChar = ' ';
			this.spnNumero.Size = new System.Drawing.Size(104, 21);
			this.spnNumero.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnNumero.TabIndex = 140;
			// 
			// spnPtoEstablec
			// 
			this.spnPtoEstablec.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsNumeracion, "FacturaSri.PuntoEmision"));
			this.spnPtoEstablec.Enabled = false;
			this.spnPtoEstablec.Location = new System.Drawing.Point(624, 64);
			this.spnPtoEstablec.MaskInput = "nnn";
			this.spnPtoEstablec.MaxValue = 999;
			this.spnPtoEstablec.MinValue = 1;
			this.spnPtoEstablec.Name = "spnPtoEstablec";
			this.spnPtoEstablec.PromptChar = ' ';
			this.spnPtoEstablec.Size = new System.Drawing.Size(104, 21);
			this.spnPtoEstablec.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnPtoEstablec.TabIndex = 139;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(504, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(120, 24);
			this.label2.TabIndex = 138;
			this.label2.Text = "Nro. Establecimientos";
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(624, 128);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(104, 23);
			this.btGenerar.TabIndex = 137;
			this.btGenerar.Text = "&Genera";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// txtCodEstab
			// 
			this.txtCodEstab.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsNumeracion, "FacturaSri.Autorizacion"));
			this.txtCodEstab.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodEstab.Enabled = false;
			this.txtCodEstab.Location = new System.Drawing.Point(624, 32);
			this.txtCodEstab.Name = "txtCodEstab";
			this.txtCodEstab.Size = new System.Drawing.Size(104, 21);
			this.txtCodEstab.TabIndex = 143;
			// 
			// spnPtoEmision
			// 
			this.spnPtoEmision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsNumeracion, "FacturaSri.NroEstabl"));
			this.spnPtoEmision.Enabled = false;
			this.spnPtoEmision.Location = new System.Drawing.Point(624, 32);
			this.spnPtoEmision.MaskInput = "nnn";
			this.spnPtoEmision.MaxValue = 999;
			this.spnPtoEmision.MinValue = 1;
			this.spnPtoEmision.Name = "spnPtoEmision";
			this.spnPtoEmision.NullText = "0";
			this.spnPtoEmision.PromptChar = ' ';
			this.spnPtoEmision.Size = new System.Drawing.Size(104, 21);
			this.spnPtoEmision.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnPtoEmision.TabIndex = 143;
			// 
			// chAprobado
			// 
			this.chAprobado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsNumeracion, "FacturaSri.Aprobado"));
			this.chAprobado.Location = new System.Drawing.Point(160, 128);
			this.chAprobado.Name = "chAprobado";
			this.chAprobado.Size = new System.Drawing.Size(128, 20);
			this.chAprobado.TabIndex = 144;
			this.chAprobado.Text = "Aprobar";
			// 
			// CompraNumeroSri
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(752, 429);
			this.Controls.Add(this.chAprobado);
			this.Controls.Add(this.spnPtoEmision);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.spnNumero);
			this.Controls.Add(this.spnPtoEstablec);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.cmbProy);
			this.Controls.Add(this.cmbCentroC);
			this.Controls.Add(this.cmbDoc);
			this.Controls.Add(this.cmdTipoF);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.txtIdFacturaSri);
			this.Controls.Add(this.txtAutorizaSri);
			this.Controls.Add(this.cmbTipoTramite);
			this.Controls.Add(this.txtActivacionAutoImpresor);
			this.Controls.Add(this.btAutoImpresor);
			this.Controls.Add(this.lblAutorizacion);
			this.Controls.Add(this.barraDato1);
			this.MaximizeBox = false;
			this.Name = "CompraNumeroSri";
			this.Text = "Compra Numercion Sri";
			this.Load += new System.EventHandler(this.CompraNumeroSri_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsNumeracion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipofactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtActivacionAutoImpresor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoTramite)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAutoImpresor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizaSri)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFacturaSri)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdTipoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProy)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPtoEstablec)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodEstab)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnPtoEmision)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
//			int IdFacturaSri = -1;
//			if (txtIdFacturaSri.Value != DBNull.Value && txtIdFacturaSri.Value != null)
//				IdFacturaSri = int.Parse(txtIdFacturaSri.Value.ToString());
//			C1.Data.FilterCondition fcFactNumero = new 
//				C1.Data.FilterCondition(cdsFacturaNumero, "FacturaNumero", "idFacturaSri = " + IdFacturaSri.ToString());
//			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
//			fcTotal.Add(fcFactNumero);
//
//			cdsFacturaNumero.Fill(fcTotal, false);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			string stExec = "Exec SriAutoImpresorCuenta '" + txtAutorizaSri.Text + "'";
			int iCuenta = Funcion.iEscalarSQL(cdsAutoImpresor, stExec, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede borrar la numeracion del punto de venta cuando es Autoimpresor",
					"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Borrar el Registro?", "Confirmacion",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stBorrarAutorizacion = "Delete FacturaNumero Where Autorizacion = '"
				+ txtAutorizaSri.Text + "'";
			Funcion.EjecutaSQL(cdsAutoImpresor, stBorrarAutorizacion);
			barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			spnPtoEmision.Value = 1;
			spnPtoEstablec.Value = 1;
			spnNumero.Value = 1000;
			chv1.Checked = false;
			chv2.Checked = false;
			chv3.Checked = false;
			chv4.Checked = false;
			chv5.Checked = false;
			barraDato1.CrearRegistro();
			
			ModificaGrilla(true);
			
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			BorrarErrores();
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if(chAprobado.Checked == true)
			{
			MessageBox.Show("No puede editar autoimpresores aprobados","Informacion");
			return;
			}
			barraDato1.EditarRegistro();
			ModificaGrilla(true);
			//			if (chkAutoImpresor.Checked)
//			{
//				txtAutorizaSri.Enabled = false;
//			
//			}
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			bool bNuevo = barraDato1.bNuevo;
			barraDato1.GrabaRegistro();
//			if (bNuevo)
			{
				string stSelect = "Exec SriActualizaAutoImpresor '" + txtIdFacturaSri.Text + "'";
				Funcion.EjecutaSQL(cdsAutoImpresor, stSelect, true);
				barraDato1.ProximoId(5);
			}
			ModificaGrilla(false);
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
//			if (chAprobado.Checked);
			string stCodTipoTra = cmbTipoTramite.Text.Trim();
			string stSelect = "Select Ruc From Seteo";
			string stRuc = Funcion.sEscalarSQL(cdsNumeracion, stSelect, true);
			stSelect = Funcion.VerificaCedula(stRuc, false);
			if (stSelect.Trim().Length > 0)
			{
				MessageBox.Show("RUC " + stRuc + " no valido", "Informacion");
				return;
			}
			string stNumAut = txtAutorizaSri.Text;
			DateTime dtFecha = DateTime.Today;
			StreamWriter sr = null;
			string stArchivo = "";
			if (stCodTipoTra == "6") stArchivo = "SolicitudAutorizacion";
			if (stCodTipoTra == "7") stArchivo = "SolicitudCambioSoftware";
			if (stCodTipoTra == "8") stArchivo = "SolicitudRenovacion";
			if (stCodTipoTra == "9") stArchivo = "SolicitudBaja";
			if (stCodTipoTra == "10") stArchivo = "SolicitudInclusion";
			if (stCodTipoTra == "11") stArchivo = "SolicitudExclusion";
			stArchivo += dtFecha.ToString("MMMyyyy") + ".XML"; 
			try
			{
				sr= File.CreateText(stArchivo);
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo " + stArchivo);
				Cursor = Cursors.Default;
				return;
			}

			sr.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n");
			sr.WriteLine("<autorizacion>=\"http://www.w3\"\n");
			sr.WriteLine("<codTipoTra>" + stCodTipoTra.ToString() + "</codTipoTra>\n");
			sr.WriteLine("<ruc>" + stRuc + "</ruc>\n");
			sr.WriteLine("<numAut>" + stNumAut + "</numAut>\n");
			sr.WriteLine("<fecha>" + dtFecha.ToString("dd/MM/yyyy") + "</fecha>\n");

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				sr.WriteLine("<detalles>\n");
				//			sr.WriteLine("\t<detalle>\n");
				//				if (dr.Cells["TipoComprobante"].Value != System.DBNull.Value)
				//					dDebe += dr.Cells["TipoComprobante"].Value.ToString();
				string stCodDoc = dr.Cells["idDocumento"].Value.ToString();
				string stEstablecimiento = dr.Cells["Establecimiento"].Text.ToString();
				string stPuntoEmision = dr.Cells["PuntoEmision"].Text.ToString();
				string stInicio = dr.Cells["Inicio"].Text.ToString();
				string stFin = dr.Cells["Fin"].Text.ToString();
				sr.WriteLine("\t<codDoc>" + stCodDoc + "</codDoc>\n");
				sr.WriteLine("\t<estab>" + stEstablecimiento + "</estab>\n");
				sr.WriteLine("\t<ptoEmi>" + stPuntoEmision + "</ptoEmi>\n");
				if (stCodTipoTra == "6" || stCodTipoTra == "10") 
					sr.WriteLine("\t<inicio>" + stInicio + "</inicio>\n");
				if (stCodTipoTra == "7" || stCodTipoTra == "8")
				{
					sr.WriteLine("\t<finOld>" + stFin + "</finOld>\n");
					sr.WriteLine("\t<iniNew>" + stInicio + "</iniNew>\n");
				}
				if (stCodTipoTra == "9" || stCodTipoTra == "11") 
					sr.WriteLine("\t<fin>" + stFin + "</fin>\n");
				sr.WriteLine("</detalles>\n");
			}
			
			sr.WriteLine("</autorizacion>\n");
			sr.Close();
			stSelect = "Exec SriImprimeAutoImpresor '" + txtIdFacturaSri.Text + "'";
			Funcion.EjecutaSQL(cdsAutoImpresor, stSelect, true);

			MessageBox.Show("Archivo " + stArchivo + " generado", "Informacion");
		}

		private void BorrarErrores()
		{
			errorProvider.SetError(ultraGrid1, "");
			//			errorProvider.SetError(txtEstacion, "");
//			errorProvider.SetError(spnBodega, "");
//			errorProvider.SetError(txtEstablecimiento, "");
//			errorProvider.SetError(txtPuntoEmision, "");
//			errorProvider.SetError(txtCodigo, "");
//			errorProvider.SetError(txtNumeracion, "");
//			errorProvider.SetError(txtInicio, "");
//			errorProvider.SetError(cmbComprobante, "");
//			errorProvider.SetError(cmbDocumento, "");
			errorProvider.SetError(cmbTipoTramite, "");
			errorProvider.SetError(txtAutorizaSri, "");
		}

		private bool Verifica()
		{
			BorrarErrores();
			bool bOk = true;
			if (txtAutorizaSri.Text.Trim().Length != 10)
			{
				errorProvider.SetError(txtAutorizaSri, "Ingrese Autorizacion del SRI");
				bOk = false;
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dr.Cells["idFacturaSri"].Value = int.Parse(txtIdFacturaSri.Value.ToString());
				if (dr.Cells["Estacion"].Value == DBNull.Value || dr.Cells["Estacion"].Value.ToString().Trim().Length == 0)
				{
					errorProvider.SetError(ultraGrid1, "Ingrese estacion");
					bOk = false;
				}
				if (dr.Cells["Bodega"].Value == DBNull.Value || (int) dr.Cells["Bodega"].Value < 0)
				{
					errorProvider.SetError(ultraGrid1, "Ingrese Bodega");
					bOk = false;
				}
				if (dr.Cells["idTipoFactura"].Value == DBNull.Value || (int) dr.Cells["idTipoFactura"].Value < 0)
				{
					errorProvider.SetError(ultraGrid1, "Ingrese Tipo de Comprobante");
					bOk = false;
				}
				if (dr.Cells["idDocumento"].Value == DBNull.Value || (int) dr.Cells["idDocumento"].Value < 0)
				{
					errorProvider.SetError(ultraGrid1, "Ingrese Tipo de Documento");
					bOk = false;
				}
				if (dr.Cells["Establecimiento"].Value == DBNull.Value || dr.Cells["Establecimiento"].Value.ToString().Trim().Length != 3)
				{
					errorProvider.SetError(ultraGrid1, "Ingrese Establecimiento de 3 digitos");
					bOk = false;
				}
				if (dr.Cells["PuntoEmision"].Value == DBNull.Value || dr.Cells["PuntoEmision"].Value.ToString().Trim().Length != 3)
				{
					errorProvider.SetError(ultraGrid1, "Ingrese Punto de Emision de 3 digitos");
					bOk = false;
				}
				if (dr.Cells["Inicio"].Value == DBNull.Value || dr.Cells["Inicio"].Value.ToString().Trim().Length < 1)
				{
					errorProvider.SetError(ultraGrid1, "Ingrese Numero inicial");
					bOk = false;
				}
				if (bOk==false) return false;
			}

			if (cmbTipoTramite.Value == DBNull.Value || cmbTipoTramite.Value.ToString().Trim().Length < 1)
			{
				errorProvider.SetError(cmbTipoTramite, "Ingrese Tipo de Tramite");
				bOk = false;
			}
		
			return bOk;
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			chv1.Enabled = bModifica;
			chv2.Enabled = bModifica;
			chv3.Enabled = bModifica;
			chv4.Enabled = bModifica;
			chv5.Enabled = bModifica;
			
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
				chAprobado.Checked = false;

			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}
//		bool AutoImpresor = false;
		private void CompraNumeroSri_Load(object sender, System.EventArgs e)
		{
			cmbFecha.Value = DateTime.Today.AddYears(1);
			barraDato1.ProximoId(4);
			barraDato1.HabilitaControles(false);
			txtIdFacturaSri.Width = 0;
			btAutoImpresor.Width=0;
			txtActivacionAutoImpresor.Width=0;
			ModificaGrilla(false);
//			string stSelect = "Select IsNull(AutoImpresor, 0) From Seteo";
//			AutoImpresor = Funcion.bEjecutaSQL(cdsNumeracion, stSelect);
//			if (!AutoImpresor)
//			{
//				txtAutorizaSri.Width = 0;
//				lblAutorizacion.Width = 0;
//			}
		}

		private void cdsNumeracion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsNumeracion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Habilita(bool bHabilita)
		{
//			cmbDocumento.Enabled = bHabilita;
			cmbTipoTramite.Enabled = bHabilita;
			txtAutorizaSri.Enabled = bHabilita;
//			chkAutoImpresor.Enabled = bHabilita;
//			txtInicio.Enabled = bHabilita;
//			txtFin.Enabled = bHabilita;
//			txtEstablecimiento.Enabled = bHabilita;
//			txtPuntoEmision.Enabled = bHabilita;
//			txtNumeracion.Enabled = !bHabilita;
//			txtCodigo.Enabled = !bHabilita;
//			chkAutomatico.Enabled = !bHabilita;
		}

		private void btAutoImpresor_Click(object sender, System.EventArgs e)
		{
			if (!barraDato1.bEditar) return;
			string stClave = txtActivacionAutoImpresor.Text.Trim();
			if (Funcion.CreaHash(stClave) != "kBXbdkEja/fTtJPo2xpowctPY4k=")
			{
				MessageBox.Show("Clave de activacion erronea", "Informacion");
				return;
			}
			Habilita(true);
		}

		private void btEstacion_Click(object sender, System.EventArgs e)
		{
			string stSelect = "Select Host_Name()";
			string stEstacion = Funcion.sEscalarSQL(cdsNumeracion, stSelect, true);
//			txtEstacion.Text = stEstacion;
//			txtEstacion.Select();
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idFacturaSri"].Value = int.Parse(txtIdFacturaSri.Value.ToString());
		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			if (this.ultraGrid1.ActiveRow.Cells["idFacturaNombre"].Value == null
				|| this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Linea", "Informacion");
				return;
			}
			Funcion.AuditoriaImprime(52, (int) ultraGrid1.ActiveRow.Cells["idFacturaNombre"].Value , true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if ( this.ultraGrid1.ActiveRow == null ||
				this.ultraGrid1.ActiveRow.Cells["idFacturaNombre"].Value == null
				)
			{
				MessageBox.Show("Seleccione Linea", "Informacion");
				return;
			}
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(52, (int) ultraGrid1.ActiveRow.Cells["idFacturaNombre"].Value, false);
			Cursor = Cursors.Default;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtIdFacturaSri_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
//		int Ichv1 = 0;
//		int Ichv2 = 0;
//		int Ichv3 = 0;
//		int Ichv4 = 0;
//		int Ichv5 = 0;
		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if(chAprobado.Checked) return;
//			if(chv1.Checked == true) Ichv1 = 1;
//			if(chv2.Checked == true) Ichv2 = 1;
//			if(chv3.Checked == true) Ichv3 = 1;
//			if(chv4.Checked == true) Ichv4 = 1;
//			if(chv5.Checked == true) Ichv5 = 1;
			int puntoemision = (int)spnPtoEmision.Value;
			int factnumero = (int)spnNumero.Value;
			int establecimiento = (int)spnPtoEstablec.Value;
			int idsri = Convert.ToInt32(txtIdFacturaSri.Value.ToString());
			if (idsri < 0 ) 
			{
				MessageBox.Show("Debe primero Guardar el registro","Informacion");
				return;
			}
			string stAuditaE = "Exec AutoimpresorGenera "+idsri;
			Funcion.EjecutaSQL(cdsAutoImpresor, stAuditaE, true);
			barraDato1.ProximoId(5);
			return;

		}

		private void cmbTipoTramite_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	

	}
}
