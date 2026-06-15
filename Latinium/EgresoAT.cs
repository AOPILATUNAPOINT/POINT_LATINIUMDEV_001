using System;
using System.Data;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Egreso.
	/// </summary>
	public class EgresoAT : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label21;
		private DevExpress.XtraEditors.CheckEdit checkEdit1;
		private C1.Data.C1DataSet cdsRetencion;
		private C1.Data.C1DataView cdvRetIva;
		private C1.Data.C1DataView cdvRetRenta;
		private C1.Data.C1DataView cdvRetTipo;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private C1.Data.C1DataView cdvCuentaCorr;
		private DevExpress.XtraEditors.LookUpEdit cmbDestino;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private System.Windows.Forms.Label label22;
		private DevExpress.XtraEditors.SimpleButton btAsiento;
		private DevExpress.XtraEditors.SimpleButton btImprimir;
		private DevExpress.XtraEditors.SimpleButton btCheque;
		private DevExpress.XtraEditors.SimpleButton btAvanzado;
		private C1.Data.C1DataSet cdsComprobante;
		private C1.Data.C1DataView cdvComprobante;
		public DevExpress.XtraEditors.TextEdit txtFactIva0;
		private DevExpress.XtraEditors.TextEdit txtFactIva;
		private DevExpress.XtraEditors.TextEdit txtDescuento;
		private DevExpress.XtraEditors.TextEdit txtPorcIva;
		private DevExpress.XtraEditors.TextEdit txtRenta0;
		private DevExpress.XtraEditors.TextEdit txtTotalFactura;
		private DevExpress.XtraEditors.TextEdit txtTotalPagar;
		private DevExpress.XtraEditors.MemoEdit txtNotas;
		private DevExpress.XtraEditors.TextEdit txtBeneficiario;
		private System.Windows.Forms.Label label20;
		private DevExpress.XtraEditors.TextEdit txtIva;
		private DevExpress.XtraEditors.SimpleButton btTotal;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.LookUpEdit cmbProveedor;
		private DevExpress.XtraEditors.LookUpEdit cmbComprobante;
		private DevExpress.XtraEditors.TextEdit txtCodAsiento;
		private C1.Data.C1DataSet cdsSeteoC;
		private DevExpress.XtraEditors.TextEdit txtCodRet;
		private C1.Data.C1DataSet cdsAsiento;
		private DevExpress.XtraEditors.RadioGroup rgContCredito;
		private C1.Data.C1DataSet cdsAsignaCuenta;
		private C1.Data.C1DataView cdvRetencion;
		private DevExpress.XtraEditors.SimpleButton btCuentaCorr;
		private DevExpress.XtraEditors.SimpleButton btProveedor;
		private DevExpress.XtraEditors.SimpleButton btComprobante;
		private DevExpress.XtraEditors.LookUpEdit cmbCuenta;
		private DevExpress.XtraEditors.TextEdit txtIdAsiento;
		private DevExpress.XtraEditors.TextEdit txtFactura;
		private DevExpress.XtraEditors.TextEdit txtIdRet;
		private DevExpress.XtraEditors.SimpleButton btEnlazar;
		private System.Windows.Forms.Label label23;
		private DevExpress.XtraEditors.TextEdit txtIvaRet1;
		private DevExpress.XtraEditors.LookUpEdit cmbPorcIvaRet1;
		private System.Windows.Forms.Label label26;
		private DevExpress.XtraEditors.TextEdit txtFactIva1;
		private System.Windows.Forms.Label label27;
		private DevExpress.XtraEditors.TextEdit txtIva1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.TextEdit txtIdCompra;
		private DevExpress.XtraEditors.TextEdit txtCheque;
		private C1.Data.C1DataSet cdsSeguridad;
		private System.Windows.Forms.Label label28;
		private DevExpress.XtraEditors.TextEdit txtDescIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRentaRet11;
		private Infragistics.Win.Misc.UltraButton btPorcRet;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaRet;
		private Infragistics.Win.Misc.UltraButton btPorcIva;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPorcIvaRet;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label lblAnulado;
		private C1.Data.C1DataSet cdsRetDetalle;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRetRenta;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdRetRenta;
		private C1.Data.C1DataView cdvCuentaNombre;
		private C1.Data.C1DataSet cdsArticuloSRI;
		private DevExpress.XtraEditors.TextEdit txtAutRet;
		private DevExpress.XtraEditors.TextEdit txtSerieRet;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public DevExpress.XtraEditors.TextEdit txtIdPagoLote;
		private DevExpress.XtraEditors.CheckEdit chkTarjeta;
		private DevExpress.XtraEditors.CheckEdit chkNoAplica;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public EgresoAT()
		{
			InitializeComponent();
		}

		public int IdAsiento = 0;
		public int IdFactura = 0;
		public EgresoAT(int Id_Asiento)
		{
			InitializeComponent();
			IdAsiento = Id_Asiento;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(EgresoAT));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RetencionIva", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RetDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetRenta", -1, "cmbRetRenta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbProveedor = new DevExpress.XtraEditors.LookUpEdit();
			this.cdsRetencion = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbComprobante = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvComprobante = new C1.Data.C1DataView();
			this.cdsComprobante = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtFactura = new DevExpress.XtraEditors.TextEdit();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtFactIva0 = new DevExpress.XtraEditors.TextEdit();
			this.txtFactIva = new DevExpress.XtraEditors.TextEdit();
			this.label6 = new System.Windows.Forms.Label();
			this.txtDescuento = new DevExpress.XtraEditors.TextEdit();
			this.txtPorcIva = new DevExpress.XtraEditors.TextEdit();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtRenta0 = new DevExpress.XtraEditors.TextEdit();
			this.txtTotalFactura = new DevExpress.XtraEditors.TextEdit();
			this.txtTotalPagar = new DevExpress.XtraEditors.TextEdit();
			this.label11 = new System.Windows.Forms.Label();
			this.txtNotas = new DevExpress.XtraEditors.MemoEdit();
			this.txtBeneficiario = new DevExpress.XtraEditors.TextEdit();
			this.label12 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.cmbCuenta = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCuentaCorr = new C1.Data.C1DataView();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.checkEdit1 = new DevExpress.XtraEditors.CheckEdit();
			this.rgContCredito = new DevExpress.XtraEditors.RadioGroup();
			this.cmbDestino = new DevExpress.XtraEditors.LookUpEdit();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.txtCodAsiento = new DevExpress.XtraEditors.TextEdit();
			this.txtCheque = new DevExpress.XtraEditors.TextEdit();
			this.cdvRetIva = new C1.Data.C1DataView();
			this.cdvRetRenta = new C1.Data.C1DataView();
			this.cdvRetTipo = new C1.Data.C1DataView();
			this.txtCodRet = new DevExpress.XtraEditors.TextEdit();
			this.label22 = new System.Windows.Forms.Label();
			this.btCuentaCorr = new DevExpress.XtraEditors.SimpleButton();
			this.btProveedor = new DevExpress.XtraEditors.SimpleButton();
			this.btComprobante = new DevExpress.XtraEditors.SimpleButton();
			this.btAsiento = new DevExpress.XtraEditors.SimpleButton();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.btCheque = new DevExpress.XtraEditors.SimpleButton();
			this.btAvanzado = new DevExpress.XtraEditors.SimpleButton();
			this.label20 = new System.Windows.Forms.Label();
			this.txtIva = new DevExpress.XtraEditors.TextEdit();
			this.btTotal = new DevExpress.XtraEditors.SimpleButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cdsAsignaCuenta = new C1.Data.C1DataSet();
			this.cdvRetencion = new C1.Data.C1DataView();
			this.txtIdAsiento = new DevExpress.XtraEditors.TextEdit();
			this.txtIdRet = new DevExpress.XtraEditors.TextEdit();
			this.btEnlazar = new DevExpress.XtraEditors.SimpleButton();
			this.label23 = new System.Windows.Forms.Label();
			this.txtIvaRet1 = new DevExpress.XtraEditors.TextEdit();
			this.cmbPorcIvaRet1 = new DevExpress.XtraEditors.LookUpEdit();
			this.label26 = new System.Windows.Forms.Label();
			this.txtFactIva1 = new DevExpress.XtraEditors.TextEdit();
			this.label27 = new System.Windows.Forms.Label();
			this.txtIva1 = new DevExpress.XtraEditors.TextEdit();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.txtIdCompra = new DevExpress.XtraEditors.TextEdit();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.txtDescIva0 = new DevExpress.XtraEditors.TextEdit();
			this.label28 = new System.Windows.Forms.Label();
			this.txtRentaRet11 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btPorcRet = new Infragistics.Win.Misc.UltraButton();
			this.txtIvaRet = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbPorcIvaRet = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btPorcIva = new Infragistics.Win.Misc.UltraButton();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.cdsRetDetalle = new C1.Data.C1DataSet();
			this.cmbRetRenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsArticuloSRI = new C1.Data.C1DataSet();
			this.grdRetRenta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvCuentaNombre = new C1.Data.C1DataView();
			this.txtAutRet = new DevExpress.XtraEditors.TextEdit();
			this.txtSerieRet = new DevExpress.XtraEditors.TextEdit();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdPagoLote = new DevExpress.XtraEditors.TextEdit();
			this.chkTarjeta = new DevExpress.XtraEditors.CheckEdit();
			this.chkNoAplica = new DevExpress.XtraEditors.CheckEdit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProveedor.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactIva0.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactIva.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcIva.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRenta0.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalPagar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rgContCredito.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDestino.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodAsiento.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodRet.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRet.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaRet1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactIva1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRentaRet11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutRet.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieRet.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPagoLote.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTarjeta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNoAplica.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(60, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Proveedor:";
			// 
			// cmbProveedor
			// 
			this.cmbProveedor.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idCliente"));
			this.cmbProveedor.Location = new System.Drawing.Point(104, 16);
			this.cmbProveedor.Name = "cmbProveedor";
			// 
			// cmbProveedor.Properties
			// 
			this.cmbProveedor.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.cmbProveedor.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbProveedor.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbProveedor.Properties.DataSource = this.cdvCliente;
			this.cmbProveedor.Properties.DisplayMember = "Nombre";
			this.cmbProveedor.Properties.NullText = "";
			this.cmbProveedor.Properties.PopupWidth = 400;
			this.cmbProveedor.Properties.ValueMember = "idCliente";
			this.cmbProveedor.Size = new System.Drawing.Size(160, 19);
			this.cmbProveedor.TabIndex = 0;
			this.cmbProveedor.Validated += new System.EventHandler(this.cmbProveedor_Validated);
			// 
			// cdsRetencion
			// 
			this.cdsRetencion.BindingContextCtrl = this;
			this.cdsRetencion.DataLibrary = "LibContabilidad";
			this.cdsRetencion.DataLibraryUrl = "";
			this.cdsRetencion.DataSetDef = "dsRetencion";
			this.cdsRetencion.FillOnRequest = false;
			this.cdsRetencion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRetencion.Name = "cdsRetencion";
			this.cdsRetencion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsRetencion.SchemaDef = null;
			this.cdsRetencion.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsRetencion_PositionChanged);
			this.cdsRetencion.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetencion_BeforeFill);
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
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idTipoRetencion"));
			this.cmbComprobante.Location = new System.Drawing.Point(104, 40);
			this.cmbComprobante.Name = "cmbComprobante";
			// 
			// cmbComprobante.Properties
			// 
			this.cmbComprobante.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.cmbComprobante.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbComprobante.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Comprobante", "Comprobante", 200, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 50, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbComprobante.Properties.DataSource = this.cdvComprobante;
			this.cmbComprobante.Properties.DisplayMember = "Comprobante";
			this.cmbComprobante.Properties.NullText = "";
			this.cmbComprobante.Properties.PopupWidth = 400;
			this.cmbComprobante.Properties.ValueMember = "idComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(160, 19);
			this.cmbComprobante.TabIndex = 1;
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
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(76, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Comprobante:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 64);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(61, 16);
			this.label3.TabIndex = 4;
			this.label3.Text = "Factura Nº:";
			// 
			// txtFactura
			// 
			this.txtFactura.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.NumFactura"));
			this.txtFactura.EditValue = "";
			this.txtFactura.Location = new System.Drawing.Point(104, 64);
			this.txtFactura.Name = "txtFactura";
			// 
			// txtFactura.Properties
			// 
			this.txtFactura.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtFactura.Size = new System.Drawing.Size(160, 19);
			this.txtFactura.TabIndex = 2;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(216, 96);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(22, 16);
			this.label4.TabIndex = 6;
			this.label4.Text = "IVA";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(24, 112);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(46, 16);
			this.label5.TabIndex = 7;
			this.label5.Text = "Subtotal";
			// 
			// txtFactIva0
			// 
			this.txtFactIva0.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.FacturaIva0"));
			this.txtFactIva0.EditValue = "";
			this.txtFactIva0.Location = new System.Drawing.Point(104, 112);
			this.txtFactIva0.Name = "txtFactIva0";
			// 
			// txtFactIva0.Properties
			// 
			this.txtFactIva0.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtFactIva0.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtFactIva0.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtFactIva0.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtFactIva0.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtFactIva0.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtFactIva0.Size = new System.Drawing.Size(88, 19);
			this.txtFactIva0.TabIndex = 4;
			// 
			// txtFactIva
			// 
			this.txtFactIva.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.FacturaIva12"));
			this.txtFactIva.EditValue = "";
			this.txtFactIva.Location = new System.Drawing.Point(200, 112);
			this.txtFactIva.Name = "txtFactIva";
			// 
			// txtFactIva.Properties
			// 
			this.txtFactIva.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtFactIva.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtFactIva.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtFactIva.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtFactIva.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtFactIva.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtFactIva.Size = new System.Drawing.Size(96, 19);
			this.txtFactIva.TabIndex = 5;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 136);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(61, 16);
			this.label6.TabIndex = 10;
			this.label6.Text = "Descuento:";
			// 
			// txtDescuento
			// 
			this.txtDescuento.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.DescIva"));
			this.txtDescuento.EditValue = "";
			this.txtDescuento.Location = new System.Drawing.Point(200, 136);
			this.txtDescuento.Name = "txtDescuento";
			// 
			// txtDescuento.Properties
			// 
			this.txtDescuento.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtDescuento.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtDescuento.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtDescuento.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtDescuento.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtDescuento.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtDescuento.Size = new System.Drawing.Size(96, 19);
			this.txtDescuento.TabIndex = 7;
			// 
			// txtPorcIva
			// 
			this.txtPorcIva.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.PorcIva"));
			this.txtPorcIva.EditValue = "";
			this.txtPorcIva.Location = new System.Drawing.Point(248, 88);
			this.txtPorcIva.Name = "txtPorcIva";
			// 
			// txtPorcIva.Properties
			// 
			this.txtPorcIva.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtPorcIva.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtPorcIva.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtPorcIva.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtPorcIva.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtPorcIva.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtPorcIva.Size = new System.Drawing.Size(48, 19);
			this.txtPorcIva.TabIndex = 3;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(24, 184);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(47, 16);
			this.label8.TabIndex = 14;
			this.label8.Text = "Renta 0:";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(24, 208);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(74, 16);
			this.label9.TabIndex = 15;
			this.label9.Text = "Total Factura:";
			// 
			// txtRenta0
			// 
			this.txtRenta0.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.Renta0"));
			this.txtRenta0.EditValue = "";
			this.txtRenta0.Location = new System.Drawing.Point(160, 184);
			this.txtRenta0.Name = "txtRenta0";
			// 
			// txtRenta0.Properties
			// 
			this.txtRenta0.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtRenta0.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtRenta0.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtRenta0.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtRenta0.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtRenta0.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtRenta0.Size = new System.Drawing.Size(136, 19);
			this.txtRenta0.TabIndex = 9;
			// 
			// txtTotalFactura
			// 
			this.txtTotalFactura.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.TotalFactura"));
			this.txtTotalFactura.EditValue = "";
			this.txtTotalFactura.Location = new System.Drawing.Point(160, 208);
			this.txtTotalFactura.Name = "txtTotalFactura";
			// 
			// txtTotalFactura.Properties
			// 
			this.txtTotalFactura.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtTotalFactura.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtTotalFactura.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalFactura.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtTotalFactura.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalFactura.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtTotalFactura.Size = new System.Drawing.Size(136, 19);
			this.txtTotalFactura.TabIndex = 10;
			// 
			// txtTotalPagar
			// 
			this.txtTotalPagar.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.Total"));
			this.txtTotalPagar.EditValue = "";
			this.txtTotalPagar.Location = new System.Drawing.Point(160, 232);
			this.txtTotalPagar.Name = "txtTotalPagar";
			// 
			// txtTotalPagar.Properties
			// 
			this.txtTotalPagar.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtTotalPagar.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtTotalPagar.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalPagar.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtTotalPagar.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtTotalPagar.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtTotalPagar.Size = new System.Drawing.Size(136, 19);
			this.txtTotalPagar.TabIndex = 11;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(24, 264);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(37, 16);
			this.label11.TabIndex = 20;
			this.label11.Text = "Notas:";
			// 
			// txtNotas
			// 
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.Notas"));
			this.txtNotas.EditValue = "";
			this.txtNotas.Location = new System.Drawing.Point(96, 256);
			this.txtNotas.Name = "txtNotas";
			// 
			// txtNotas.Properties
			// 
			this.txtNotas.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtNotas.Size = new System.Drawing.Size(200, 40);
			this.txtNotas.TabIndex = 12;
			// 
			// txtBeneficiario
			// 
			this.txtBeneficiario.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.Beneficiario"));
			this.txtBeneficiario.EditValue = "";
			this.txtBeneficiario.Location = new System.Drawing.Point(440, 96);
			this.txtBeneficiario.Name = "txtBeneficiario";
			// 
			// txtBeneficiario.Properties
			// 
			this.txtBeneficiario.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtBeneficiario.Size = new System.Drawing.Size(208, 19);
			this.txtBeneficiario.TabIndex = 20;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(352, 96);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(67, 16);
			this.label12.TabIndex = 22;
			this.label12.Text = "Beneficiario:";
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(352, 16);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(91, 16);
			this.label14.TabIndex = 26;
			this.label14.Text = "% Iva Ret Bienes";
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(360, 216);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(39, 16);
			this.label16.TabIndex = 28;
			this.label16.Text = "Fecha:";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(352, 120);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(44, 16);
			this.label17.TabIndex = 30;
			this.label17.Text = "Cuenta:";
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idCuentaCorriente"));
			this.cmbCuenta.Location = new System.Drawing.Point(440, 120);
			this.cmbCuenta.Name = "cmbCuenta";
			// 
			// cmbCuenta.Properties
			// 
			this.cmbCuenta.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.cmbCuenta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCuenta.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					new DevExpress.XtraEditors.Controls.LookUpColumnInfo("NumCuenta", "NumCuenta", 70, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbCuenta.Properties.DataSource = this.cdvCuentaCorr;
			this.cmbCuenta.Properties.DisplayMember = "NumCuenta";
			this.cmbCuenta.Properties.NullText = "";
			this.cmbCuenta.Properties.ValueMember = "idCuentaCorriente";
			this.cmbCuenta.Size = new System.Drawing.Size(176, 19);
			this.cmbCuenta.TabIndex = 21;
			this.cmbCuenta.Validated += new System.EventHandler(this.cmbCuenta_Validated);
			// 
			// cdvCuentaCorr
			// 
			this.cdvCuentaCorr.BindingContextCtrl = this;
			this.cdvCuentaCorr.DataSet = this.cdsCuentaCorr;
			this.cdvCuentaCorr.TableName = "";
			this.cdvCuentaCorr.TableViewName = "CuentaCorriente";
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(352, 144);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 16);
			this.label18.TabIndex = 32;
			this.label18.Text = "Cheque:";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(352, 168);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(48, 16);
			this.label19.TabIndex = 33;
			this.label19.Text = "Compra:";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.Location = new System.Drawing.Point(24, 328);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(69, 16);
			this.label21.TabIndex = 35;
			this.label21.Text = "Cod. Asiento";
			// 
			// checkEdit1
			// 
			this.checkEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.IncluirIva"));
			this.checkEdit1.Location = new System.Drawing.Point(24, 304);
			this.checkEdit1.Name = "checkEdit1";
			// 
			// checkEdit1.Properties
			// 
			this.checkEdit1.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.checkEdit1.Properties.Caption = "Incluir IVA";
			this.checkEdit1.Size = new System.Drawing.Size(96, 21);
			this.checkEdit1.TabIndex = 36;
			// 
			// rgContCredito
			// 
			this.rgContCredito.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.ContadoCredito"));
			this.rgContCredito.Location = new System.Drawing.Point(216, 304);
			this.rgContCredito.Name = "rgContCredito";
			// 
			// rgContCredito.Properties
			// 
			this.rgContCredito.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.rgContCredito.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																					new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Contado"),
																																																					new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Crédito")});
			this.rgContCredito.Size = new System.Drawing.Size(75, 48);
			this.rgContCredito.TabIndex = 34;
			// 
			// cmbDestino
			// 
			this.cmbDestino.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idCuenta"));
			this.cmbDestino.Location = new System.Drawing.Point(440, 168);
			this.cmbDestino.Name = "cmbDestino";
			// 
			// cmbDestino.Properties
			// 
			this.cmbDestino.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.cmbDestino.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDestino.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbDestino.Properties.DataSource = this.cdvCuenta;
			this.cmbDestino.Properties.DisplayMember = "CodRapido";
			this.cmbDestino.Properties.NullText = "";
			this.cmbDestino.Properties.PopupWidth = 300;
			this.cmbDestino.Properties.ValueMember = "idCuenta";
			this.cmbDestino.Size = new System.Drawing.Size(176, 19);
			this.cmbDestino.TabIndex = 23;
			this.cmbDestino.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbDestino_MouseDown);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo = false";
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
			// txtCodAsiento
			// 
			this.txtCodAsiento.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.CodAsiento"));
			this.txtCodAsiento.EditValue = "";
			this.txtCodAsiento.Location = new System.Drawing.Point(112, 328);
			this.txtCodAsiento.Name = "txtCodAsiento";
			// 
			// txtCodAsiento.Properties
			// 
			this.txtCodAsiento.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtCodAsiento.Size = new System.Drawing.Size(96, 19);
			this.txtCodAsiento.TabIndex = 33;
			// 
			// txtCheque
			// 
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.Cheque"));
			this.txtCheque.EditValue = "";
			this.txtCheque.Location = new System.Drawing.Point(440, 144);
			this.txtCheque.Name = "txtCheque";
			// 
			// txtCheque.Properties
			// 
			this.txtCheque.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtCheque.Size = new System.Drawing.Size(176, 19);
			this.txtCheque.TabIndex = 22;
			// 
			// cdvRetIva
			// 
			this.cdvRetIva.BindingContextCtrl = this;
			this.cdvRetIva.DataSet = this.cdsRetencion;
			this.cdvRetIva.TableName = "";
			this.cdvRetIva.TableViewName = "RetencionIva";
			// 
			// cdvRetRenta
			// 
			this.cdvRetRenta.BindingContextCtrl = this;
			this.cdvRetRenta.DataSet = this.cdsRetencion;
			this.cdvRetRenta.TableName = "";
			this.cdvRetRenta.TableViewName = "RetencionRenta";
			// 
			// cdvRetTipo
			// 
			this.cdvRetTipo.BindingContextCtrl = this;
			this.cdvRetTipo.DataSet = this.cdsRetencion;
			this.cdvRetTipo.TableName = "";
			this.cdvRetTipo.TableViewName = "RetencionTipo";
			// 
			// txtCodRet
			// 
			this.txtCodRet.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.Codigo"));
			this.txtCodRet.EditValue = "";
			this.txtCodRet.Location = new System.Drawing.Point(584, 72);
			this.txtCodRet.Name = "txtCodRet";
			// 
			// txtCodRet.Properties
			// 
			this.txtCodRet.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtCodRet.Size = new System.Drawing.Size(64, 19);
			this.txtCodRet.TabIndex = 19;
			this.txtCodRet.ToolTip = "Secuencial de Retencion";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Location = new System.Drawing.Point(352, 72);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(40, 16);
			this.label22.TabIndex = 47;
			this.label22.Text = "Nº Ret.";
			// 
			// btCuentaCorr
			// 
			this.btCuentaCorr.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btCuentaCorr.Location = new System.Drawing.Point(624, 120);
			this.btCuentaCorr.Name = "btCuentaCorr";
			this.btCuentaCorr.Size = new System.Drawing.Size(24, 23);
			this.btCuentaCorr.TabIndex = 21;
			this.btCuentaCorr.Text = "...";
			this.btCuentaCorr.Click += new System.EventHandler(this.btCuentaCorr_Click);
			// 
			// btProveedor
			// 
			this.btProveedor.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btProveedor.Location = new System.Drawing.Point(272, 16);
			this.btProveedor.Name = "btProveedor";
			this.btProveedor.Size = new System.Drawing.Size(24, 23);
			this.btProveedor.TabIndex = 55;
			this.btProveedor.Text = "...";
			this.btProveedor.Click += new System.EventHandler(this.btProveedor_Click);
			// 
			// btComprobante
			// 
			this.btComprobante.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btComprobante.Location = new System.Drawing.Point(272, 40);
			this.btComprobante.Name = "btComprobante";
			this.btComprobante.Size = new System.Drawing.Size(24, 23);
			this.btComprobante.TabIndex = 56;
			this.btComprobante.Text = "...";
			this.btComprobante.Click += new System.EventHandler(this.btComprobante_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btAsiento.Image = ((System.Drawing.Image)(resources.GetObject("btAsiento.Image")));
			this.btAsiento.Location = new System.Drawing.Point(328, 376);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(24, 23);
			this.btAsiento.TabIndex = 58;
			this.btAsiento.ToolTip = "Mostrar Asiento";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btImprimir
			// 
			this.btImprimir.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btImprimir.Image")));
			this.btImprimir.Location = new System.Drawing.Point(376, 376);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(24, 23);
			this.btImprimir.TabIndex = 59;
			this.btImprimir.ToolTip = "Imprime Egreso";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// btCheque
			// 
			this.btCheque.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btCheque.Image = ((System.Drawing.Image)(resources.GetObject("btCheque.Image")));
			this.btCheque.Location = new System.Drawing.Point(400, 376);
			this.btCheque.Name = "btCheque";
			this.btCheque.Size = new System.Drawing.Size(24, 23);
			this.btCheque.TabIndex = 60;
			this.btCheque.ToolTip = "Imprime Cheque";
			this.btCheque.Click += new System.EventHandler(this.btCheque_Click);
			// 
			// btAvanzado
			// 
			this.btAvanzado.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btAvanzado.Image = ((System.Drawing.Image)(resources.GetObject("btAvanzado.Image")));
			this.btAvanzado.Location = new System.Drawing.Point(424, 376);
			this.btAvanzado.Name = "btAvanzado";
			this.btAvanzado.Size = new System.Drawing.Size(24, 23);
			this.btAvanzado.TabIndex = 61;
			this.btAvanzado.Text = "Avanzado";
			this.btAvanzado.ToolTip = "Mostrar Factura";
			this.btAvanzado.Click += new System.EventHandler(this.btAvanzado_Click);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(24, 160);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(26, 16);
			this.label20.TabIndex = 62;
			this.label20.Text = "IVA:";
			// 
			// txtIva
			// 
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.IvaTotal"));
			this.txtIva.EditValue = "";
			this.txtIva.Location = new System.Drawing.Point(160, 160);
			this.txtIva.Name = "txtIva";
			// 
			// txtIva.Properties
			// 
			this.txtIva.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIva.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtIva.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIva.Properties.EditFormat.FormatString = "#,##0.00";
			this.txtIva.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIva.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtIva.Size = new System.Drawing.Size(136, 19);
			this.txtIva.TabIndex = 8;
			// 
			// btTotal
			// 
			this.btTotal.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btTotal.Location = new System.Drawing.Point(24, 232);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(88, 23);
			this.btTotal.TabIndex = 64;
			this.btTotal.Text = "Total a Pagar";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
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
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
			// 
			// cdsAsignaCuenta
			// 
			this.cdsAsignaCuenta.BindingContextCtrl = this;
			this.cdsAsignaCuenta.DataLibrary = "LibContabilidad";
			this.cdsAsignaCuenta.DataLibraryUrl = "";
			this.cdsAsignaCuenta.DataSetDef = "dsAsignaCuenta";
			this.cdsAsignaCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsignaCuenta.Name = "cdsAsignaCuenta";
			this.cdsAsignaCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsignaCuenta.SchemaDef = null;
			this.cdsAsignaCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsignaCuenta_BeforeFill);
			// 
			// cdvRetencion
			// 
			this.cdvRetencion.BindingContextCtrl = this;
			this.cdvRetencion.DataSet = this.cdsRetencion;
			this.cdvRetencion.Sort = "idRetencion";
			this.cdvRetencion.TableName = "";
			this.cdvRetencion.TableViewName = "Retencion";
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idAsiento"));
			this.txtIdAsiento.EditValue = "";
			this.txtIdAsiento.Location = new System.Drawing.Point(304, 184);
			this.txtIdAsiento.Name = "txtIdAsiento";
			// 
			// txtIdAsiento.Properties
			// 
			this.txtIdAsiento.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIdAsiento.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIdAsiento.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtIdAsiento.Size = new System.Drawing.Size(40, 19);
			this.txtIdAsiento.TabIndex = 65;
			this.txtIdAsiento.TabStop = false;
			this.txtIdAsiento.EditValueChanged += new System.EventHandler(this.txtIdAsiento_EditValueChanged);
			// 
			// txtIdRet
			// 
			this.txtIdRet.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idRetencion"));
			this.txtIdRet.EditValue = "";
			this.txtIdRet.Location = new System.Drawing.Point(304, 208);
			this.txtIdRet.Name = "txtIdRet";
			// 
			// txtIdRet.Properties
			// 
			this.txtIdRet.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIdRet.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIdRet.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtIdRet.Size = new System.Drawing.Size(40, 19);
			this.txtIdRet.TabIndex = 66;
			this.txtIdRet.TabStop = false;
			// 
			// btEnlazar
			// 
			this.btEnlazar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btEnlazar.Image = ((System.Drawing.Image)(resources.GetObject("btEnlazar.Image")));
			this.btEnlazar.Location = new System.Drawing.Point(352, 376);
			this.btEnlazar.Name = "btEnlazar";
			this.btEnlazar.Size = new System.Drawing.Size(24, 23);
			this.btEnlazar.TabIndex = 57;
			this.btEnlazar.ToolTip = "Enlazar a asiento";
			this.btEnlazar.Click += new System.EventHandler(this.btEnlazar_Click);
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.Location = new System.Drawing.Point(352, 40);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(79, 16);
			this.label23.TabIndex = 68;
			this.label23.Text = "% Iva Ret Serv";
			// 
			// txtIvaRet1
			// 
			this.txtIvaRet1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.IvaRetenido1"));
			this.txtIvaRet1.EditValue = "";
			this.txtIvaRet1.Location = new System.Drawing.Point(504, 40);
			this.txtIvaRet1.Name = "txtIvaRet1";
			// 
			// txtIvaRet1.Properties
			// 
			this.txtIvaRet1.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIvaRet1.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtIvaRet1.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIvaRet1.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIvaRet1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtIvaRet1.Size = new System.Drawing.Size(112, 19);
			this.txtIvaRet1.TabIndex = 16;
			// 
			// cmbPorcIvaRet1
			// 
			this.cmbPorcIvaRet1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idRetIva1"));
			this.cmbPorcIvaRet1.Location = new System.Drawing.Point(440, 40);
			this.cmbPorcIvaRet1.Name = "cmbPorcIvaRet1";
			// 
			// cmbPorcIvaRet1.Properties
			// 
			this.cmbPorcIvaRet1.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.cmbPorcIvaRet1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbPorcIvaRet1.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Porcentaje", "Porcentaje", 25, DevExpress.Utils.FormatType.Numeric, "", true, DevExpress.Utils.HorzAlignment.Far),
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Detalle", "Detalle", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default)});
			this.cmbPorcIvaRet1.Properties.DataSource = this.cdvRetIva;
			this.cmbPorcIvaRet1.Properties.DisplayMember = "Porcentaje";
			this.cmbPorcIvaRet1.Properties.NullText = "";
			this.cmbPorcIvaRet1.Properties.PopupWidth = 100;
			this.cmbPorcIvaRet1.Properties.ShowFooter = false;
			this.cmbPorcIvaRet1.Properties.ShowHeader = false;
			this.cmbPorcIvaRet1.Properties.ValueMember = "idRetIva";
			this.cmbPorcIvaRet1.Size = new System.Drawing.Size(56, 19);
			this.cmbPorcIvaRet1.TabIndex = 15;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.Location = new System.Drawing.Point(352, 192);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(80, 16);
			this.label26.TabIndex = 75;
			this.label26.Text = "Subtotal IVA 1:";
			// 
			// txtFactIva1
			// 
			this.txtFactIva1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.FacturaIva121"));
			this.txtFactIva1.EditValue = "";
			this.txtFactIva1.Location = new System.Drawing.Point(440, 192);
			this.txtFactIva1.Name = "txtFactIva1";
			// 
			// txtFactIva1.Properties
			// 
			this.txtFactIva1.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtFactIva1.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtFactIva1.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtFactIva1.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtFactIva1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtFactIva1.Size = new System.Drawing.Size(104, 19);
			this.txtFactIva1.TabIndex = 24;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.Location = new System.Drawing.Point(552, 192);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(26, 16);
			this.label27.TabIndex = 79;
			this.label27.Text = "IVA:";
			// 
			// txtIva1
			// 
			this.txtIva1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.IvaTotal1"));
			this.txtIva1.EditValue = "";
			this.txtIva1.Location = new System.Drawing.Point(584, 192);
			this.txtIva1.Name = "txtIva1";
			// 
			// txtIva1.Properties
			// 
			this.txtIva1.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIva1.Properties.DisplayFormat.FormatString = "#,##0.00";
			this.txtIva1.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIva1.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIva1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtIva1.Size = new System.Drawing.Size(64, 19);
			this.txtIva1.TabIndex = 25;
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btAceptar.Location = new System.Drawing.Point(24, 376);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 81;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Visible = false;
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(160, 376);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 82;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Visible = false;
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idCompra"));
			this.txtIdCompra.EditValue = "";
			this.txtIdCompra.Location = new System.Drawing.Point(304, 232);
			this.txtIdCompra.Name = "txtIdCompra";
			// 
			// txtIdCompra.Properties
			// 
			this.txtIdCompra.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIdCompra.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIdCompra.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtIdCompra.Size = new System.Drawing.Size(40, 19);
			this.txtIdCompra.TabIndex = 83;
			this.txtIdCompra.TabStop = false;
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
			// txtDescIva0
			// 
			this.txtDescIva0.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.DescIva0"));
			this.txtDescIva0.EditValue = "";
			this.txtDescIva0.Location = new System.Drawing.Point(104, 136);
			this.txtDescIva0.Name = "txtDescIva0";
			// 
			// txtDescIva0.Properties
			// 
			this.txtDescIva0.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtDescIva0.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtDescIva0.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtDescIva0.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtDescIva0.Size = new System.Drawing.Size(88, 19);
			this.txtDescIva0.TabIndex = 6;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.Location = new System.Drawing.Point(160, 96);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(35, 16);
			this.label28.TabIndex = 85;
			this.label28.Text = "IVA 0:";
			// 
			// txtRentaRet11
			// 
			this.txtRentaRet11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRetencion, "Retencion.RentaRetenida"));
			this.txtRentaRet11.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRentaRet11.FormatString = "#,##0.00";
			this.txtRentaRet11.Location = new System.Drawing.Point(528, 216);
			this.txtRentaRet11.Name = "txtRentaRet11";
			this.txtRentaRet11.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRentaRet11.PromptChar = ' ';
			this.txtRentaRet11.Size = new System.Drawing.Size(80, 21);
			this.txtRentaRet11.TabIndex = 27;
			// 
			// btPorcRet
			// 
			this.btPorcRet.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPorcRet.Location = new System.Drawing.Point(616, 216);
			this.btPorcRet.Name = "btPorcRet";
			this.btPorcRet.Size = new System.Drawing.Size(24, 23);
			this.btPorcRet.TabIndex = 88;
			this.btPorcRet.Text = "...";
			this.btPorcRet.Click += new System.EventHandler(this.btPorcRet_Click);
			// 
			// txtIvaRet
			// 
			this.txtIvaRet.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRetencion, "Retencion.IvaRetenido"));
			this.txtIvaRet.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaRet.FormatString = "#,##0.00";
			this.txtIvaRet.Location = new System.Drawing.Point(504, 16);
			this.txtIvaRet.Name = "txtIvaRet";
			this.txtIvaRet.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaRet.PromptChar = ' ';
			this.txtIvaRet.Size = new System.Drawing.Size(112, 21);
			this.txtIvaRet.TabIndex = 14;
			// 
			// cmbPorcIvaRet
			// 
			this.cmbPorcIvaRet.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPorcIvaRet.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPorcIvaRet.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRetencion, "Retencion.idRetIva"));
			this.cmbPorcIvaRet.DataSource = this.cdvRetIva;
			ultraGridBand1.AddButtonCaption = "RetencionRenta";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 210;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbPorcIvaRet.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPorcIvaRet.DisplayMember = "Porcentaje";
			this.cmbPorcIvaRet.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPorcIvaRet.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPorcIvaRet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPorcIvaRet.Location = new System.Drawing.Point(440, 16);
			this.cmbPorcIvaRet.Name = "cmbPorcIvaRet";
			this.cmbPorcIvaRet.Size = new System.Drawing.Size(56, 21);
			this.cmbPorcIvaRet.TabIndex = 13;
			this.cmbPorcIvaRet.ValueMember = "idRetIva";
			// 
			// btPorcIva
			// 
			this.btPorcIva.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPorcIva.Location = new System.Drawing.Point(624, 16);
			this.btPorcIva.Name = "btPorcIva";
			this.btPorcIva.Size = new System.Drawing.Size(24, 23);
			this.btPorcIva.TabIndex = 91;
			this.btPorcIva.Text = "...";
			this.btPorcIva.Click += new System.EventHandler(this.btPorcIva_Click);
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Retencion";
			this.barraDato1.DataNombreId = "idRetencion";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsRetencion;
			this.barraDato1.DataTabla = "Retencion";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 376);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 98;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// lblAnulado
			// 
			this.lblAnulado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblAnulado.Location = new System.Drawing.Point(520, 352);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(120, 23);
			this.lblAnulado.TabIndex = 99;
			this.lblAnulado.Text = "ANULADO";
			// 
			// cdsRetDetalle
			// 
			this.cdsRetDetalle.BindingContextCtrl = this;
			this.cdsRetDetalle.DataLibrary = "LibContabilidad";
			this.cdsRetDetalle.DataLibraryUrl = "";
			this.cdsRetDetalle.DataSetDef = "dsRetDetalle";
			this.cdsRetDetalle.FillOnRequest = false;
			this.cdsRetDetalle.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRetDetalle.Name = "cdsRetDetalle";
			this.cdsRetDetalle.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsRetDetalle.SchemaDef = null;
			this.cdsRetDetalle.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetDetalle_BeforeFill);
			// 
			// cmbRetRenta
			// 
			this.cmbRetRenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRetRenta.DataMember = "ArticuloSRI";
			this.cmbRetRenta.DataSource = this.cdsArticuloSRI;
			ultraGridBand2.AddButtonCaption = "RetencionRenta";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 69;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 59;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbRetRenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbRetRenta.DisplayMember = "Porcentaje";
			this.cmbRetRenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRetRenta.Location = new System.Drawing.Point(368, 256);
			this.cmbRetRenta.Name = "cmbRetRenta";
			this.cmbRetRenta.Size = new System.Drawing.Size(256, 64);
			this.cmbRetRenta.TabIndex = 101;
			this.cmbRetRenta.ValueMember = "idArticuloSRI";
			this.cmbRetRenta.Visible = false;
			// 
			// cdsArticuloSRI
			// 
			this.cdsArticuloSRI.BindingContextCtrl = this;
			this.cdsArticuloSRI.DataLibrary = "LibFacturacion";
			this.cdsArticuloSRI.DataLibraryUrl = "";
			this.cdsArticuloSRI.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloSRI.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloSRI.Name = "cdsArticuloSRI";
			this.cdsArticuloSRI.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloSRI.SchemaDef = null;
			// 
			// grdRetRenta
			// 
			this.grdRetRenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdRetRenta.DataMember = "RetDetalle";
			this.grdRetRenta.DataSource = this.cdsRetDetalle;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdRetRenta.DisplayLayout.Appearance = appearance1;
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance2;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 1;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance3;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.Caption = "Tipo Ret.";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn18.Width = 81;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.grdRetRenta.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.grdRetRenta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdRetRenta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdRetRenta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.grdRetRenta.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdRetRenta.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdRetRenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdRetRenta.Location = new System.Drawing.Point(360, 240);
			this.grdRetRenta.Name = "grdRetRenta";
			this.grdRetRenta.Size = new System.Drawing.Size(288, 112);
			this.grdRetRenta.TabIndex = 28;
			this.grdRetRenta.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.grdRetRenta_AfterRowInsert);
			this.grdRetRenta.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdRetRenta_AfterCellUpdate);
			// 
			// cdvCuentaNombre
			// 
			this.cdvCuentaNombre.BindingContextCtrl = this;
			this.cdvCuentaNombre.DataSet = this.cdsCuenta;
			this.cdvCuentaNombre.Sort = "Codigo";
			this.cdvCuentaNombre.TableName = "";
			this.cdvCuentaNombre.TableViewName = "Cuenta";
			// 
			// txtAutRet
			// 
			this.txtAutRet.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.AutorizaRet"));
			this.txtAutRet.EditValue = "";
			this.txtAutRet.Location = new System.Drawing.Point(440, 72);
			this.txtAutRet.Name = "txtAutRet";
			// 
			// txtAutRet.Properties
			// 
			this.txtAutRet.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtAutRet.Size = new System.Drawing.Size(72, 19);
			this.txtAutRet.TabIndex = 17;
			this.txtAutRet.ToolTip = "Autorizacion de Retencion";
			// 
			// txtSerieRet
			// 
			this.txtSerieRet.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.SerieRet"));
			this.txtSerieRet.EditValue = "";
			this.txtSerieRet.Location = new System.Drawing.Point(520, 72);
			this.txtSerieRet.Name = "txtSerieRet";
			// 
			// txtSerieRet.Properties
			// 
			this.txtSerieRet.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtSerieRet.Size = new System.Drawing.Size(56, 19);
			this.txtSerieRet.TabIndex = 18;
			this.txtSerieRet.ToolTip = "Serie de Retencion";
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			// 
			// dtFecha
			// 
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRetencion, "Retencion.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(408, 216);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.TabIndex = 26;
			this.dtFecha.Value = new System.DateTime(2006, 9, 12, 0, 0, 0, 0);
			// 
			// txtIdPagoLote
			// 
			this.txtIdPagoLote.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.idPagoLote"));
			this.txtIdPagoLote.EditValue = "";
			this.txtIdPagoLote.Location = new System.Drawing.Point(304, 256);
			this.txtIdPagoLote.Name = "txtIdPagoLote";
			// 
			// txtIdPagoLote.Properties
			// 
			this.txtIdPagoLote.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.txtIdPagoLote.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtIdPagoLote.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtIdPagoLote.Size = new System.Drawing.Size(40, 19);
			this.txtIdPagoLote.TabIndex = 106;
			this.txtIdPagoLote.TabStop = false;
			// 
			// chkTarjeta
			// 
			this.chkTarjeta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.Tarjeta"));
			this.chkTarjeta.Location = new System.Drawing.Point(472, 376);
			this.chkTarjeta.Name = "chkTarjeta";
			// 
			// chkTarjeta.Properties
			// 
			this.chkTarjeta.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkTarjeta.Properties.Caption = "Tarjeta";
			this.chkTarjeta.Size = new System.Drawing.Size(64, 21);
			this.chkTarjeta.TabIndex = 107;
			// 
			// chkNoAplica
			// 
			this.chkNoAplica.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsRetencion, "Retencion.NoAplica"));
			this.chkNoAplica.Location = new System.Drawing.Point(560, 376);
			this.chkNoAplica.Name = "chkNoAplica";
			// 
			// chkNoAplica.Properties
			// 
			this.chkNoAplica.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkNoAplica.Properties.Caption = "No Aplica";
			this.chkNoAplica.Size = new System.Drawing.Size(80, 21);
			this.chkNoAplica.TabIndex = 108;
			// 
			// EgresoAT
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(672, 414);
			this.Controls.Add(this.chkNoAplica);
			this.Controls.Add(this.chkTarjeta);
			this.Controls.Add(this.txtIdPagoLote);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtSerieRet);
			this.Controls.Add(this.txtAutRet);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbRetRenta);
			this.Controls.Add(this.lblAnulado);
			this.Controls.Add(this.btPorcIva);
			this.Controls.Add(this.cmbPorcIvaRet);
			this.Controls.Add(this.txtIvaRet);
			this.Controls.Add(this.btPorcRet);
			this.Controls.Add(this.txtRentaRet11);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDescIva0);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.txtIva1);
			this.Controls.Add(this.txtFactIva1);
			this.Controls.Add(this.txtIvaRet1);
			this.Controls.Add(this.cmbPorcIvaRet1);
			this.Controls.Add(this.txtIdRet);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.btAvanzado);
			this.Controls.Add(this.btCheque);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btEnlazar);
			this.Controls.Add(this.btComprobante);
			this.Controls.Add(this.btProveedor);
			this.Controls.Add(this.btCuentaCorr);
			this.Controls.Add(this.txtCodRet);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.txtCodAsiento);
			this.Controls.Add(this.cmbDestino);
			this.Controls.Add(this.rgContCredito);
			this.Controls.Add(this.checkEdit1);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.txtBeneficiario);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtTotalPagar);
			this.Controls.Add(this.txtTotalFactura);
			this.Controls.Add(this.txtRenta0);
			this.Controls.Add(this.txtPorcIva);
			this.Controls.Add(this.txtDescuento);
			this.Controls.Add(this.txtFactIva);
			this.Controls.Add(this.txtFactIva0);
			this.Controls.Add(this.txtFactura);
			this.Controls.Add(this.cmbComprobante);
			this.Controls.Add(this.cmbProveedor);
			this.Controls.Add(this.grdRetRenta);
			this.Controls.Add(this.barraDato1);
			this.MaximizeBox = false;
			this.Name = "EgresoAT";
			this.Text = "Egreso";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.EgresoAT_Closing);
			this.Load += new System.EventHandler(this.Egreso_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbProveedor.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactIva0.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactIva.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcIva.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRenta0.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalPagar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rgContCredito.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDestino.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodAsiento.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodRet.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRet.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaRet1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactIva1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRentaRet11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPorcIvaRet)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutRet.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieRet.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPagoLote.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkTarjeta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkNoAplica.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		SeteoCRow_tableView drSeteoC;
		C1.Data.C1DataRow drSeteo;
		public C1.Data.C1DataRow drFactura;
		private Acceso miAcceso;
		public double fEgreso = 0;
		private void Egreso_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			miAcceso =new Acceso("030401", this.cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			this.barraDato1.stFiltroExt = " And IngresoEgreso = 2";

			this.txtIdAsiento.Width =0;
			this.txtIdRet.Width = 0;
			this.txtIdCompra.Width = 0;
			txtIdPagoLote.Width = 0;
			C1DataRow dr = this.cdsSeteoC.TableViews[0].Rows[0];
			drSeteoC = SeteoCRow_tableView.Obj(dr);
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];

			if (IdAsiento > 0) // Pantalla pedida desde Asiento
			{
				string stSelect = "Select idCompra From Compra Where idAsiento = " 
					+ IdAsiento.ToString();
				int idF = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
				if (idF > 0) // Asiento tiene una factura asociada
				{
					stSelect = "Select idCompra From Retencion Where idCompra = " + idF.ToString();
					idF = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
					if (idF > 0) // Asiento tiene una factura asociada
					{
						IdAsiento = 0;
						IdFactura = idF;
					}
				}
			}

			if (IdFactura > 0)
			{
				string stFiltro = " And idCompra = " + IdFactura.ToString().Trim();
				this.barraDato1.BuscaRegistro(barraDato1.stFiltroExt + stFiltro);
				if (this.barraDato1.IdRegistro <= 0)
				{
					int IdPagoLote = 0;
					if (txtIdPagoLote.EditValue != System.DBNull.Value 
						&& txtIdPagoLote.EditValue.ToString().Length > 0)
						IdPagoLote = int.Parse(txtIdPagoLote.EditValue.ToString());
					string stSelect = "EXEC RetencionCrea " + IdFactura.ToString();
					Funcion.EjecutaSQL(cdsRetDetalle, stSelect, true);
					this.barraDato1.BuscaRegistro(barraDato1.stFiltroExt + stFiltro);
					if (drSeteo["SriSerieRet"] != System.DBNull.Value)
						drRet.SerieRet = drSeteo["SriSerieRet"].ToString();
					if (drSeteo["SriAutRet"] != System.DBNull.Value)
						drRet.AutorizaRet = drSeteo["SriAutRet"].ToString();
					drRet.idPagoLote = IdPagoLote;

					// Habilita Controles
					this.barraDato1.Visible = false;
					this.btAceptar.Visible = false;
					this.btAsiento.Visible = false;
					this.btAvanzado.Visible = false;
					this.btEnlazar.Visible = false;
					this.btImprimir.Visible = false;
					this.btCheque.Visible = false;

					this.btAceptar.Visible = true;
					this.btCancelar.Visible = true;
				} // Fin de Factura no encontrada
				else // Factura Encontrada
				{
					this.barraDato1.Edicion(false);
				} // Fin de factura encontrada
				return;
			} // Fin de Factura

			
			if (IdAsiento > 0) // Pantalla pedida desde Asiento
			{
				string stFiltro = "idAsiento = " + IdAsiento.ToString().Trim();
				string stSelec = "Select idRetencion FROM Retencion Where " + stFiltro;
				int IdRet = Funcion.iEscalarSQL(cdsRetencion, stSelec, false);
//				DataTable dt = this.cdsRetencion.StorageDataSet.Tables["Retencion"];
	//			DataRow [] dr1 = dt.Select(stFiltro);
				if (IdRet < 1) //(dr1.Length == 0)
				{
					IdFactura = -1; // Simula que viene de una factura para no generar el asiento
					this.cdvRetencion.Sort = "idRetencion";
					this.barraDato1.CrearRegistro();
					ValoresInciales();
					drRet.idAsiento = IdAsiento;
					drRet.idCliente = 0;
					drRet.Fecha = (DateTime) this.drFactura["Fecha"];
					drRet.Total = fEgreso;
					drRet.CodAsiento = (string) this.drFactura["CodAsiento"];

					drRet.idTipoRetencion = 0;
					drRet.NumFactura = "";
					drRet.FacturaIva12 = 0;
					drRet.FacturaIva0 = fEgreso;
					drRet.DescIva = 0;
					drRet.DescIva0 = 0;
					drRet.ContadoCredito = 1;
					drRet.PorcIva = drSeteoC.Iva;
					drRet.IngresoEgreso = 2;
					drRet.AutorizaRet = drSeteo["SriAutRet"].ToString();
					drRet.SerieRet = drSeteo["SriSerieRet"].ToString();

					// Habilita Controles
					this.barraDato1.Visible = false;
					this.btAceptar.Visible = false;
					this.btAsiento.Visible = false;
					this.btAvanzado.Visible = false;
					this.btEnlazar.Visible = false;
					this.btImprimir.Visible = false;
					this.btCheque.Visible = false;

					this.btAceptar.Visible = true;
					this.btCancelar.Visible = true;
					return;
				}
				else
				{
					this.BuscarRetencion(IdRet);
					this.barraDato1.Edicion(false);
				}
			}
			else
			{
				if (IdAsiento < 0)
					this.BuscarRetencion(-IdAsiento);
				else
					this.barraDato1.PosUltima();
			}
			this.barraDato1.HabilitaControles(false);
		}

		private void btPorcRet_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("502", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Tabla de Porcentajes de Retención en Renta",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			using (ArticuloSRI miArt = new ArticuloSRI())
			{
				if (miArt.ShowDialog() == DialogResult.OK)
				{
					cdsArticuloSRI.Clear();
					cdsArticuloSRI.Fill();
				}
			}
		}

		private void btPorcIva_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("503", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Tabla de Porcentajes de Retención del IVA",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			using (RetIVA miRetIva = new RetIVA())
			{
				if (miRetIva.ShowDialog() == DialogResult.OK)
				{
					this.cdsRetencion.Clear();
					this.cdsRetencion.Fill();
				}
			}
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a la Eliminación de Retenciones",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.txtIdRet.EditValue.ToString().Trim().Length == 0)
			{
				MessageBox.Show("No existen retenciones para Borrar",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (drRet.Borrar == false) 
			{
				if (DialogResult.No ==	MessageBox.Show(
					"¿Está seguro que desea Anular la Retención?", 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
					MessageBoxDefaultButton.Button2)) return;
				drRet.Borrar = true;
				drRet.RentaRetenida = 0;
				drRet.RentaRetenida1 = 0;
				drRet.IvaRetenido = 0;
				drRet.IvaRetenido1 = 0;
				cdsRetencion.Update();
				string stSelect = "Update RetDetalle Set Retenido = 0 "
					+ "Where idRetencion = " + this.txtIdRet.Text;
				Funcion.EjecutaSQL(cdsRetDetalle, stSelect, true);
				barraDato1.ProximoId(5);
			}
			else
			{
				if (DialogResult.No ==	MessageBox.Show(
					"¿Está seguro que desea borrar la Retención y el Asiento correspondiente?", 
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
					MessageBoxDefaultButton.Button2)) return;
				this.barraDato1.BorraRegistro();
			}
		}

		private void BuscarRetencion(int id)
		{
			this.barraDato1.IdRegistro = id;
			this.barraDato1.ProximoId(5);
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Buscar)
			{
				MessageBox.Show("No tiene acceso a Busqueda de Egresos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			using (BuscaRetencion miBusqueda = new BuscaRetencion(2))
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					try
					{
						int id = int.Parse(miBusqueda.idBusca.Text);
						this.BuscarRetencion(id);
					}
					catch{}
				}
			}
		}

		private void ValoresInciales()
		{
			drRet.Fecha = DateTime.Today;
			drRet.PorcIva = drSeteoC.Iva;
			drRet.idCuenta = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,14);
			int iNumCuentaCorr = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,10);
			drRet.idCuentaCorriente = Contabilidad.NumeroCuentaCorr(this.cdsAsignaCuenta,iNumCuentaCorr);
			if (drRet.idCuentaCorriente > 0)
			{
				int iNumCheque = Contabilidad.NumeroCheque(this.cdsCuentaCorr, drRet.idCuentaCorriente, false);
				if (iNumCheque  != -1)
					drRet.Cheque = iNumCheque.ToString();
			}

			drRet.idRetIva1 = 1;
			drRet.idRetRenta1 = 1;
			drRet.idRetIva = 1;
			drRet.idRetRenta = 1;
			drRet.idTipoRetencion = 1;
			drRet.CodigoSri = "325";
			drRet.ContadoCredito = 1;
			drRet.IngresoEgreso = 2;
			drRet.SerieRet = drSeteo["SriSerieRet"].ToString();
			drRet.AutorizaRet = drSeteo["SriAutRet"].ToString();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a la Creación de Egresos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.barraDato1.CrearRegistro();
			this.ValoresInciales();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.LimpiaErrores();
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Edición de Egresos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (this.txtIdRet.EditValue.ToString().Trim().Length == 0)
			{
				MessageBox.Show("No existen retenciones para Editar",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.EditarRegistro();
			drRet = RetencionRow_tableView.Obj(
				this.cdsRetencion.TableViews["Retencion"].Rows.Find((int) this.txtIdRet.EditValue));
		}

		private bool Grabacion()
		{
			if (cmbPorcIvaRet.Value != null && (int) cmbPorcIvaRet.Value > 0 && cmbPorcIvaRet.Text.Length > 0 
				&& (int.Parse(cmbPorcIvaRet.Text) > 0 || int.Parse(cmbPorcIvaRet1.Text) > 0))
			{
				if (cmbProveedor.EditValue != null)
				{
					int IdCliente = (int) cmbProveedor.EditValue;
					string stCliente = "Select IsNull(ContribEspecial, 0) From Cliente Where idCliente = "
						+ IdCliente.ToString();
					bool bEspecial = Funcion.bEscalarSQL(cdsCliente, stCliente, false);
					if (bEspecial) 
					{
						MessageBox.Show("No puede Haber retenciones de Iva en Contribuyentes Especiales", 
							"Informacion", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}
				}
			}
			this.txtBeneficiario.Focus();
			this.Total();
			this.Refresh();
			if (!this.Verifica()) return false;
			if (IdFactura == 0) this.CreaAsiento();
			//if (IdFactura == -1) CreaCheque();

			double dRetendido = double.Parse(txtRentaRet11.Value.ToString())
				+ double.Parse(txtIvaRet.Value.ToString());
			if ((btAceptar.Visible == true) || (barraDato1.bNuevo &&  dRetendido > 0.001) )
			{
				// Numeracion de Retencion
				string stCodRet = Contabilidad.NumeroAsiento(this.cdsAsiento, 7);
				if(stCodRet.Trim().Length > 0) drRet.Codigo = stCodRet;
			}
			else
			{
				drRet.Editar = true;
				drRet.FechaEditado = DateTime.Now;
			}
			int iNumRet = int.Parse(this.txtIdRet.Text);
			//Eliminar estas lineas
			string stFecha = "Update Retencion Set Fecha = '"
				+ drRet.Fecha.ToString("yyyyMMdd") + "' Where idRetencion = "
				+ drRet.idRetencion.ToString();
			Funcion.EjecutaSQL(cdsRetencion, stFecha, true);

			this.barraDato1.GrabaRegistro();
			if (IdFactura > 0)
			{
				string stSelect = "Update Compra Set Retencion = 1 Where IdCompra = "
					+ IdFactura.ToString() + " And Retencion <> 1";
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			}
			// Actualiza las Retencions de la Renta
			try
			{
				cdsRetDetalle.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Informacion");
			}
			int iNumRetN = int.Parse(this.txtIdRet.Text);
			string stSelect1 = "Update RetDetalle Set idRetencion = " +
				iNumRetN.ToString() + " Where idRetencion = " + iNumRet.ToString();
			Funcion.EjecutaSQL(cdsRetencion, stSelect1, true);
			return true;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.Grabacion();
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a la Impresion de Retenciones",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (txtIdRet.EditValue.ToString().Trim().Length == 0)
			{
				MessageBox.Show("No existen retenciones para Imprimir",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Retencion.idRetencion} = " + txtIdRet.EditValue;
			Reporte miReporte = new Reporte("RetencionMarzo.rpt", stFiltro);
			miReporte.Show();
			Cursor = Cursors.Arrow;
		}

		private void Total()
		{
			double fBaseRenta = 0;
			double fBaseRenta1 = 0;
			double fBaseIVA = 0;
			double fPorcRentaRet = 0;
			double fPorcIvaRet = 0;

			fBaseIVA = (drRet.FacturaIva12 - drRet.DescIva) * drRet.PorcIva;
			drRet.IvaTotal = (double) Math.Round(fBaseIVA / 100.0, iNumDec);
			fBaseIVA = drRet.FacturaIva121 * drRet.PorcIva;
			drRet.IvaTotal1 = (double) Math.Round(fBaseIVA / 100.0, iNumDec);

			//Calculo de Renta
			fBaseRenta = drRet.FacturaIva0 + drRet.FacturaIva12 - drRet.DescIva 
				- drRet.DescIva0;
			fBaseRenta1 = drRet.FacturaIva01 + drRet.FacturaIva121;
			//Total de Factura
			drRet.TotalFactura = fBaseRenta + fBaseRenta1 + 
				drRet.IvaTotal1 + drRet.IvaTotal + drRet.Renta0;

			//Calculo de retenciones
			this.cmbPorcIvaRet.Select();
			this.cmbPorcIvaRet1.Select();
			if (cmbPorcIvaRet.Text.Length > 0)
				fPorcIvaRet = double.Parse(cmbPorcIvaRet.Text.ToString());
			drRet.IvaRetenido = (double) Math.Round(fPorcIvaRet * drRet.IvaTotal / 100.0, iNumDec);
			if (cmbPorcIvaRet1.Text.Length > 0)
				fPorcIvaRet = double.Parse(cmbPorcIvaRet1.Text.ToString());
			drRet.IvaRetenido1 = (double) Math.Round(fPorcIvaRet * drRet.IvaTotal1 / 100.0, iNumDec);


			fBaseRenta = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdRetRenta.Rows.All)
			{
				fBaseRenta += (double) dr.Cells["Retenido"].Value;
			}
			drRet.RentaRetenida = (double) Math.Round(fBaseRenta, iNumDec);
			fBaseRenta1 = fBaseRenta1 * fPorcRentaRet;
			drRet.RentaRetenida1 = (double) Math.Round(fBaseRenta1 / 100.0, iNumDec);
			//Total a pagar 
			drRet.Total = drRet.TotalFactura - drRet.IvaRetenido - drRet.RentaRetenida
				- drRet.IvaRetenido1 - drRet.RentaRetenida1;
		}

		public RetencionRow_tableView drRet;
		private void cdsRetencion_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "Retencion")
				drRet = RetencionRow_tableView.Obj(e.Row);
		}

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			this.Total();
		}


		private void LimpiaErrores()
		{
			this.errorProvider.SetError(this.cmbProveedor, "");
			this.errorProvider.SetError(this.txtBeneficiario, "");
			this.errorProvider.SetError(this.txtTotalPagar, "");
			this.errorProvider.SetError(this.txtCodAsiento, "");
			this.errorProvider.SetError(this.txtCodRet, "");
			this.errorProvider.SetError(this.dtFecha, "");
			this.errorProvider.SetError(this.txtFactura, "");
			this.errorProvider.SetError(this.cmbCuenta, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.LimpiaErrores();
//			if (drRet.NumFactura.Trim().Length == 0) 
//			{
//				bOk = false;
//				this.errorProvider.SetError(this.txtFactura, "Ingrese Número de Factura");
//			}
			if (drRet.IsidCuentaCorrienteNull() || drRet.idCuentaCorriente == 0) 
			{
				if ((int) this.rgContCredito.EditValue == 1 && IdFactura <= 0) // Contado obliga a tener cuenta
				{
					bOk = false;
					errorProvider.SetError(cmbCuenta, "Ingrese Cuenta Corriente");
				}
			}

			if (drRet.IsidClienteNull()) 
			{
				bOk = false;
				this.errorProvider.SetError(this.cmbProveedor, "Ingrese Proveedor");
			}

			if (drRet.Beneficiario.Trim().Length == 0) 
			{
				bOk = false;
				this.errorProvider.SetError(this.txtBeneficiario, "Ingrese Beneficiario");
			}

			if (drRet.Total == 0) {
				bOk = false;
				this.errorProvider.SetError(this.txtTotalPagar, "Valor a pagar erroneo");
			}

      
			// Verificacion de Fecha
			if (drRet.IsFechaNull()) 
			{
					bOk = false;
					this.errorProvider.SetError(this.dtFecha, "Debe Ingresar Fecha de Retención");
			}
			else // Existe Fecha
			{
				if (drSeteoC.IsFechaFinNull() || drSeteoC.IsFechaIniNull())
				{
					this.errorProvider.SetError(this.dtFecha, "");
				}
				else // Existen las tres fechas
				{
					if (drRet.Fecha > drSeteoC.FechaFin || drRet.Fecha < drSeteoC.FechaIni)
					{
						bOk = false;
						this.errorProvider.SetError(this.dtFecha, "Fecha debe estar entre el " 
							+ drSeteoC.FechaIni.ToString("dd/MMM/yyyy")+ " y el " 
							+ drSeteoC.FechaFin.ToString("dd/MMM/yyyy") + ".\n"
							+ "Fechas definidas en Archivo - Propiedades - Contabilidad");
					}
				}
			}
			// Fin de verificacion de fecha


			return bOk;
		}

		private void btAvanzado_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("2", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Pantalla de Compras",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.txtIdCompra.EditValue == System.DBNull.Value || 
				this.txtIdCompra.EditValue.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Esta retención no fue generada desde una factura",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.Cursor = Cursors.WaitCursor;
			Compra miCompra = new Compra(1, (int) this.txtIdCompra.EditValue);
			miCompra.MdiParent = this.MdiParent;
			miCompra.Show();
			this.Cursor = Cursors.Arrow;
		}

		private void NumeraAsiento()
		{
			// Objeto de numeracion de asientos
			// Contado = 1 Credito = 2
			int iTipoAsi = 3;
			if ((int) this.rgContCredito.EditValue == 2) // Contado
				iTipoAsi = 1;
			// Numeracion de Asiento
			string stCodAsi = Contabilidad.NumeroAsiento(this.cdsAsiento, iTipoAsi);
			if(stCodAsi.Trim().Length > 0) drRet.CodAsiento = stCodAsi;
		}

		private int iNumDec = 2;
		private void CreaDetalleAsi(int IdAsiento, int IdCuenta, double fDebe, double fHaber)
		{
			// Creacion de Detalle de Asiento
			C1DataRow c1drDetAsi = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			AsientoDetalleRow_tableView drDetAsi = AsientoDetalleRow_tableView.Obj(c1drDetAsi);
			drDetAsi.idAsiento = IdAsiento;
			drDetAsi.idCuenta = IdCuenta;
			drDetAsi.Debe = Math.Round(fDebe, iNumDec);
			drDetAsi.Haber = Math.Round(fHaber, iNumDec);
		}

		// Cuando viene de Asientos
		private void CreaCheque()
		{
			C1DataRow drAsi;
			drAsi = this.cdsAsiento.TableViews["Asiento"].Rows.Find(drRet.idAsiento);
			int iNumCheque = Contabilidad.NumeroCheque(this.cdsCuentaCorr, (int) this.cmbCuenta.EditValue, true);
			if (iNumCheque != -1)	drAsi["Cheque"] = iNumCheque.ToString();
			else drAsi["Cheque"] = this.txtCheque.EditValue;
			this.cdsAsiento.Update();
		}
		private void CreaAsiento()
		{
			C1DataRow c1drAsi;
			AsientoRow_tableView drAsi;
			if (drRet.idAsiento > 0)
			{
				if (DialogResult.No == MessageBox.Show("Asiento ya existe desea reemplazarlo?"
					, "Confirmación",MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
				c1drAsi = this.cdsAsiento.TableViews["Asiento"].Rows.Find(drRet.idAsiento);
				if (c1drAsi == null) 
				{
						MessageBox.Show("Error IdAsiento " + drRet.idAsiento.ToString() 
							+ " no existe");
					return;
				}
				C1DataRow[] drDetalle = c1drAsi.GetChildRows();
				if (drDetalle.Length > 0)
					foreach(C1DataRow drDet in drDetalle)
						drDet.Delete();

				drAsi = AsientoRow_tableView.Obj(c1drAsi);
			}
			else
			{
				// Creacion de un nuevo asiento
				c1drAsi = this.cdsAsiento.TableViews["Asiento"].AddNew();
				drAsi = AsientoRow_tableView.Obj(c1drAsi);
				this.NumeraAsiento();
				int iNumCheque = Contabilidad.NumeroCheque(this.cdsCuentaCorr, (int) this.cmbCuenta.EditValue, true);
				if (iNumCheque != -1)	drRet.Cheque = iNumCheque.ToString();
				else 
				{
					if (txtCheque.Text == "-1") drRet.Cheque = "";
				}
			}

			drAsi.Descripcion = drRet.Beneficiario;
			if ((int) this.rgContCredito.EditValue == 1) // Contado
				drAsi.idTipoAsiento = 3; // Egreso
			else // Credito
				drAsi.idTipoAsiento = 1; // Asiento
			drAsi.CodAsiento = (string) drRet.CodAsiento;
			drAsi.Cheque = drRet.Cheque;
			drAsi.Fecha = drRet.Fecha;
			drAsi.Notas = drRet.Notas;
			drAsi.Numero = drRet.Codigo; // Numero de retencion
			drAsi.Total = drRet.Total + drRet.RentaRetenida + drRet.IvaRetenido;
			drAsi.TotalH = drAsi.Total;
			this.cdsAsiento.Update();


			// Añade linea de mercaderia
			double fValor = 0;
			int IdCuenta = 0;
			int IdAsiento = drAsi.idAsiento;
			if (drRet.IncluirIva)
				fValor = drRet.TotalFactura + drRet.DescIva + drRet.DescIva0;
			else
			{
				fValor = drRet.TotalFactura + drRet.DescIva 
					+ drRet.DescIva0 - drRet.IvaTotal - drRet.IvaTotal1;
			}
			this.CreaDetalleAsi(IdAsiento, drRet.idCuenta, fValor, 0);

			// Añade Valor de IVA si incluye iva
			if (drRet.IvaTotal + drRet.IvaTotal1 > 0 && !drRet.IncluirIva)
			{
				IdCuenta = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta, 11); // Iva en compras
				fValor = drRet.IvaTotal + drRet.IvaTotal1;
				this.CreaDetalleAsi(IdAsiento, IdCuenta, fValor, 0);
			}

			// Retencion de IVA
			if (drRet.IvaRetenido > 0)
			{
				// Linea de retencion de IVA
				IdCuenta = AsignaRetIva();
				fValor = drRet.IvaRetenido;
				this.CreaDetalleAsi(IdAsiento, IdCuenta, 0, fValor);
			}

			// Retencion de IVA1
			if (drRet.IvaRetenido1 > 0)
			{
				// Linea de retencion de IVA
				IdCuenta = AsignaRetIva1();
				fValor = drRet.IvaRetenido1;
				this.CreaDetalleAsi(IdAsiento, IdCuenta, 0, fValor);
			}

			// Retencion de Renta
			if (drRet.RentaRetenida > 0)
			{
				// Linea de retencion de renta
				IdCuenta = AsignaRetRenta();
				fValor = drRet.RentaRetenida;
				this.CreaDetalleAsi(IdAsiento, IdCuenta, 0, fValor);
			}

			// Retencion de Renta
			if (drRet.RentaRetenida1 > 0)
			{
				// Linea de retencion de renta
				IdCuenta = AsignaRetRenta1();
				fValor = drRet.RentaRetenida1;
				this.CreaDetalleAsi(IdAsiento, IdCuenta, 0, fValor);
			}

			// Caja o cuenta por pagar
			if (drRet.Total > 0)
			{
				// Linea de retencion de caja o cuenta por pagar
				if (drRet.ContadoCredito == 1) // Contado
				{
					if (this.cmbCuenta.EditValue != System.DBNull.Value)
						IdCuenta = Contabilidad.CuentaCuentaCorr(this.cdsCuentaCorr, (int) this.cmbCuenta.EditValue);
					else
						IdCuenta = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,10);
				}
				else // Cuentas por pagar
					IdCuenta = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,13);
				fValor = drRet.Total;
      
				this.CreaDetalleAsi(IdAsiento, IdCuenta, 0, fValor);
			}

			// Descuentos
			if (drRet.DescIva > 0)
			{
				IdCuenta = Contabilidad.AsignaCuenta(this.cdsAsignaCuenta,12);
				fValor = drRet.DescIva;
				this.CreaDetalleAsi(IdAsiento, IdCuenta, 0, fValor);
			}
			this.cdsAsiento.Update();
			if (this.barraDato1.bNuevo)
				drRet.idAsiento = drAsi.idAsiento;
		}


		private int AsignaRetIva()
		{
			C1DataRow dr = cdsRetencion.TableViews["RetencionIva"].Rows[drRet.idRetIva-1];
			return (int) dr["idCuentaE"];
		}

		private int AsignaRetRenta()
		{
			C1DataRow dr = cdsRetencion.TableViews["RetencionRenta"].Rows[drRet.idRetRenta-1];
			return (int) dr["idCuentaE"];
		}

		private int AsignaRetIva1()
		{
			C1DataRow dr = cdsRetencion.TableViews["RetencionIva"].Rows[drRet.idRetIva1-1];
			return (int) dr["idCuentaE"];
		}

		private int AsignaRetRenta1()
		{
			C1DataRow dr = cdsRetencion.TableViews["RetencionRenta"].Rows[drRet.idRetRenta1-1];
			return (int) dr["idCuentaE"];
		}


		private void btEnlazar_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("138", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Pantalla de Asientos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.txtIdAsiento.EditValue.ToString().Trim().Length == 0) 
			{
				MessageBox.Show("No Existe Asiento",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.Cursor = Cursors.WaitCursor;
			try
			{
				miAsiento.BuscaAsiento((int) this.txtIdAsiento.EditValue);
				miAsiento.Show();
			}
			catch
			{
				MessageBox.Show(
					"Debe Abrir la pantalla Asiento, con el botón Anterior Mostrar Asiento \npara aplicar esta opción.",
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			this.Cursor = Cursors.Arrow;
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("139", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Impresión de Egresos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (this.txtIdAsiento.EditValue.ToString().Trim().Length == 0)
			{
				MessageBox.Show("No existe Egreso para Imprimir",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.Cursor = Cursors.WaitCursor;
			string stFiltro = "{Asiento.idAsiento} = " + this.txtIdAsiento.EditValue;
			string stReporte = "";
			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			if ((int) this.rgContCredito.EditValue == 1)
			{
				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "ValorCheque";
				discreteVal.Value = float.Parse(this.txtTotalPagar.Text);
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);
				stReporte = "Egreso.Rpt";
			}
			else
				stReporte = "Diario.Rpt";
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			if ((int) this.rgContCredito.EditValue == 1)
				miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();		
			this.Cursor = Cursors.Arrow;
		}

		private void btCheque_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("361", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Impresion de Cheques",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (this.txtIdAsiento.EditValue.ToString().Trim().Length == 0)
			{
				MessageBox.Show("No existe Cheque para Imprimir",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.Cursor = Cursors.WaitCursor;
			string stFiltro = "{Asiento.idAsiento} = " + this.txtIdAsiento.EditValue;
			string stCheque;
			string stSelect = "SELECT Reporte FROM CuentaCorriente " +
				"WHERE idCuentaCorriente = " + cmbCuenta.EditValue.ToString();
			stCheque = Funcion.sEscalarSQL(cdsCuenta, stSelect, false);
			if (stCheque.Length == 0) stCheque = "Cheque.rpt";

			Reporte miReporte = new Reporte(stCheque, stFiltro);

			// Definicion de Parametros
			double fValorCheque = (double) this.txtTotalPagar.EditValue;
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "ValorCheque";
			discreteVal.Value = fValorCheque;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			// Fin de definicion de parametros
			
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();
			this.Cursor = Cursors.Arrow;
		}

		private void btProveedor_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("71", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la pantalla de Proveedores",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			Cliente miCliente;
			
			if (this.cmbProveedor.EditValue == System.DBNull.Value || this.cmbProveedor.EditValue == null)
				miCliente = new Cliente(false);
			else
				miCliente = new Cliente(false, (int) this.cmbProveedor.EditValue);
			miCliente.ShowDialog();
			miCliente.Dispose();
			miCliente = null;
			this.cdsCliente.Clear();
			this.cdsCliente.Fill();
			this.cmbCuenta.Refresh();
		}

		private void btCuentaCorr_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("275", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a Cuentas Corrientes",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			using (CuentaCorriente miCuentaCorr = new CuentaCorriente())
			{
				miCuentaCorr.ShowDialog();
			}
			this.cdsCuentaCorr.Clear();
			this.cdsCuentaCorr.Fill();
			this.cmbCuenta.Refresh();
		}

		Asiento miAsiento;
		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("138", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Pantalla de Asientos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (this.txtIdAsiento.EditValue.ToString().Trim().Length == 0) 
			{
				MessageBox.Show("No Existe Asiento",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.Cursor = Cursors.WaitCursor;
			miAsiento = new Asiento((int) this.txtIdAsiento.EditValue);
			miAsiento.MdiParent = this.MdiParent;

			miAsiento.Show();
			this.Cursor = Cursors.Arrow;
		}

		private void cmbProveedor_Validated(object sender, System.EventArgs e)
		{
			if (this.cmbProveedor.EditValue != null)
			{
				drRet.Beneficiario = cmbProveedor.Text.Trim();
				string stSelect = "Select idRetIva From Cliente Where idCliente= "
					+ cmbProveedor.EditValue.ToString();
				int iRet = Funcion.iEscalarSQL(cdsCliente, stSelect, false);
				if (iRet > 0)	drRet.idRetIva = Funcion.iEscalarSQL(cdsCliente, stSelect);
				
//				stSelect = "Select idRetRenta From Cliente Where idCliente= "
//					+ cmbProveedor.EditValue.ToString();
//				iRet = Funcion.iEscalarSQL(cdsCliente, stSelect, false);
//				if (iRet > 0)	drRet.idRetRenta = Funcion.iEscalarSQL(cdsCliente, stSelect);
			}
			else
				drRet.Beneficiario = "";
		}

		private void cdsRetencion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRetencion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsComprobante_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsComprobante.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCliente"];

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Ambos] <> 0 OR [Proveedor] <> 0"));
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsAsignaCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsignaCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!this.Grabacion()) return;
			barraDato1.Visible = true;
			this.DialogResult = DialogResult.OK;
		}

		private void cmbCuenta_Validated(object sender, System.EventArgs e)
		{
			if (! this.barraDato1.bNuevo) return; // Solo numera en nuevos registros
			if (this.cmbCuenta.EditValue == System.DBNull.Value) return;
			int iCheque = Contabilidad.NumeroCheque(this.cdsCuentaCorr, (int) this.cmbCuenta.EditValue, false);
			if (iCheque != -1) drRet.Cheque = iCheque.ToString();
		}

		private void btComprobante_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("505", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Tabla de Tipos de Comprobantes",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			RetComprobante miComprobante = new RetComprobante();
			miComprobante.MdiParent= this.MdiParent;
			miComprobante.Show();
		}

		private void btCodigoSri_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("504", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a la Tabla de Códigos de SRI",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			RetCodigoSRI miCodigoSri = new RetCodigoSRI();
			miCodigoSri.MdiParent= this.MdiParent;
			miCodigoSri.Show();		
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			if (!drRet.Borrar) return;
			if (DialogResult.No ==	MessageBox.Show(
				"¿Está seguro que desea Restaurar la Retención?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Information, 
				MessageBoxDefaultButton.Button2)) return;

			drRet.Borrar = false;
			btTotal_Click(this, e);
			drRet.Borrar = false;
			cdsRetencion.Update();
			string stSelect = "Update RetDetalle set RetDetalle.Retenido = "
				+ "IsNull(articulosri.porcentaje, 0) * IsNull(retdetalle.base, 0) / 100.0 "
				+ "from RetDetalle Inner Join ArticuloSri On "
				+ "RetDetalle.idRetRenta = ArticuloSri.idArticuloSri "
				+ "Where RetDetalle.idRetencion = " + txtIdRet.Text;
			Funcion.EjecutaSQL(cdsRetDetalle, stSelect, true);

			barraDato1.ProximoId(5);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drRet = RetencionRow_tableView.Obj(
					this.cdsRetencion.TableViews["Retencion"].Rows.Find((int) this.txtIdRet.EditValue));
				if (drRet.Borrar) 
					lblAnulado.Visible = true;
				else
					lblAnulado.Visible = false;

				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(this.cdsRetDetalle, "RetDetalle", 
						"idRetencion = " + txtIdRet.Text);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				this.cdsRetDetalle.Fill(fcTotal, false);
			}
			catch(System.Exception ex)
			{
//				MessageBox.Show(ex.Message, "Error en Egreso");
			}
		}

		private void grdRetRenta_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idRetencion"].Value = int.Parse(this.txtIdRet.Text);
		}

		private void grdRetRenta_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "Base" || e.Cell.Column.ToString() == "idRetRenta")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						double dBase = 0;
						int IdArtSRI = 0;
						if (e.Cell.Row.Cells["Base"].Value != System.DBNull.Value)
							dBase = (double) e.Cell.Row.Cells["Base"].Value;
						if (e.Cell.Row.Cells["idRetRenta"].Value != System.DBNull.Value)
							IdArtSRI = (int) e.Cell.Row.Cells["idRetRenta"].Value;
						string stSelect = "Select Porcentaje From ArticuloSRI Where "
							+ "idArticuloSRI = " + IdArtSRI.ToString();
						double dPorcentaje = Funcion.dEscalarSQL(cdsRetencion, stSelect, false);
						double dValorRet = Math.Round(dBase* dPorcentaje/100.0 + .0001, 2);
						e.Cell.Row.Cells["Retenido"].Value = dValorRet;
					}
		}

		private void cmbDestino_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (MouseButtons.Right != e.Button) return;
			if (this.cmbDestino.Properties.DisplayMember == "CodRapido")
			{
				this.cmbDestino.Properties.Columns.Clear();
				this.cmbDestino.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
				this.cmbDestino.Properties.DataSource = this.cdvCuentaNombre;
				this.cmbDestino.Properties.DisplayMember = "Descripcion";
			}
			else
			{
				this.cmbDestino.Properties.Columns.Clear();
				this.cmbDestino.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("CodRapido", "CodRapido", 66, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near),																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Descripcion", "Descripcion", 68, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
				this.cmbDestino.Properties.DataSource = this.cdvCuenta;
				this.cmbDestino.Properties.DisplayMember = "CodRapido";
			}
		}

		private void cdsRetDetalle_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRetDetalle.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void EgresoAT_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!barraDato1.Visible)
			{
				try
				{
					string stSelect = "Delete RetDetalle Where idRetencion = "
						+ drRet.idRetencion.ToString();
					Funcion.EjecutaSQL(cdsAsiento, stSelect, true);
					stSelect = "Delete Retencion Where idRetencion = "
						+ drRet.idRetencion.ToString();
					Funcion.EjecutaSQL(cdsAsiento, stSelect, true);
					barraDato1.Visible = true;
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Informacion");
				}
			}
		}

		private void txtIdAsiento_EditValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
