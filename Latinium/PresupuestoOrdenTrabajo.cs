using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for PresupuestoOrdenTrabajo.
	/// </summary>
	public class PresupuestoOrdenTrabajo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label10;
		private C1.Data.C1DataSet cdsSolicitudPago;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCapitulo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPago;
		private C1.Data.C1DataSet cdsPago;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPagoFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorPagos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMultas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorIva;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnticipos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFondoG;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdSolicitud;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetIva;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetRenta;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorPagar;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFacturado;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private System.Windows.Forms.Label label17;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PresupuestoOrdenTrabajo()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_P_PresupuestoDetSolicitudPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetSolicituPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCapitulo", -1, "cmbCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodFideicomiso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSolicitudPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescripcionCapitulo", 0);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			this.cmbPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSolicitudPago = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtPagoFecha = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtValorPagos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAnticipos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMultas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtValorIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFondoG = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbCapitulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsPago = new C1.Data.C1DataSet();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtIdSolicitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRetIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtRetRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtValorPagar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtFacturado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSolicitudPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagoFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMultas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapitulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorPagar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacturado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbPago
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPago.Appearance = appearance1;
			this.cmbPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.idPago"));
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 8;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 152;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbPago.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPago.DisplayMember = "";
			this.cmbPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPago.Location = new System.Drawing.Point(107, 8);
			this.cmbPago.Name = "cmbPago";
			this.cmbPago.Size = new System.Drawing.Size(93, 21);
			this.cmbPago.TabIndex = 153;
			this.cmbPago.ValueMember = "";
			this.cmbPago.Validated += new System.EventHandler(this.cmbPago_Validated);
			// 
			// cdsSolicitudPago
			// 
			this.cdsSolicitudPago.BindingContextCtrl = this;
			this.cdsSolicitudPago.DataLibrary = "LibFacturacion";
			this.cdsSolicitudPago.DataLibraryUrl = "";
			this.cdsSolicitudPago.DataSetDef = "dsPresupuestoC";
			this.cdsSolicitudPago.FillOnRequest = false;
			this.cdsSolicitudPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSolicitudPago.Name = "cdsSolicitudPago";
			this.cdsSolicitudPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSolicitudPago.SchemaDef = null;
			this.cdsSolicitudPago.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsSolicitudPago_PositionChanged);
			this.cdsSolicitudPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsSolicitudPago_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(7, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 16);
			this.label1.TabIndex = 154;
			this.label1.Text = "Solicitud del Pago:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(7, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(59, 16);
			this.label2.TabIndex = 157;
			this.label2.Text = "Proveedor:";
			// 
			// txtNotas
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance2;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.Notas"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(88, 280);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(336, 40);
			this.txtNotas.TabIndex = 158;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(7, 64);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 159;
			this.label3.Text = "Facturas:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(216, 96);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(87, 16);
			this.label4.TabIndex = 160;
			this.label4.Text = "Pago a la Fecha:";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(208, 120);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(109, 16);
			this.label5.TabIndex = 161;
			this.label5.Text = "Valor Presente Pago:";
			// 
			// txtPagoFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPagoFecha.Appearance = appearance3;
			this.txtPagoFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.PagoFecha"));
			this.txtPagoFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPagoFecha.FormatString = "#,##0.00";
			this.txtPagoFecha.Location = new System.Drawing.Point(320, 96);
			this.txtPagoFecha.Name = "txtPagoFecha";
			this.txtPagoFecha.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPagoFecha.PromptChar = ' ';
			this.txtPagoFecha.ReadOnly = true;
			this.txtPagoFecha.Size = new System.Drawing.Size(104, 22);
			this.txtPagoFecha.TabIndex = 162;
			// 
			// txtValorPagos
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorPagos.Appearance = appearance4;
			this.txtValorPagos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.ValorPresente"));
			this.txtValorPagos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorPagos.FormatString = "#,##0.00";
			this.txtValorPagos.Location = new System.Drawing.Point(320, 120);
			this.txtValorPagos.Name = "txtValorPagos";
			this.txtValorPagos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorPagos.PromptChar = ' ';
			this.txtValorPagos.ReadOnly = true;
			this.txtValorPagos.Size = new System.Drawing.Size(104, 22);
			this.txtValorPagos.TabIndex = 163;
			// 
			// txtAnticipos
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnticipos.Appearance = appearance5;
			this.txtAnticipos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.Anticipo"));
			this.txtAnticipos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnticipos.FormatString = "#,##0.00";
			this.txtAnticipos.Location = new System.Drawing.Point(320, 168);
			this.txtAnticipos.Name = "txtAnticipos";
			this.txtAnticipos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAnticipos.PromptChar = ' ';
			this.txtAnticipos.ReadOnly = true;
			this.txtAnticipos.Size = new System.Drawing.Size(104, 22);
			this.txtAnticipos.TabIndex = 167;
			// 
			// txtMultas
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMultas.Appearance = appearance6;
			this.txtMultas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.Multas"));
			this.txtMultas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMultas.FormatString = "#,##0.00";
			this.txtMultas.Location = new System.Drawing.Point(320, 144);
			this.txtMultas.Name = "txtMultas";
			this.txtMultas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMultas.PromptChar = ' ';
			this.txtMultas.ReadOnly = true;
			this.txtMultas.Size = new System.Drawing.Size(104, 22);
			this.txtMultas.TabIndex = 166;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(216, 176);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(53, 16);
			this.label6.TabIndex = 165;
			this.label6.Text = "Anticipos:";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(216, 152);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(41, 16);
			this.label7.TabIndex = 164;
			this.label7.Text = "Multas:";
			// 
			// txtValorIva
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorIva.Appearance = appearance7;
			this.txtValorIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.Iva"));
			this.txtValorIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorIva.FormatString = "#,##0.00";
			this.txtValorIva.Location = new System.Drawing.Point(104, 144);
			this.txtValorIva.Name = "txtValorIva";
			this.txtValorIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorIva.PromptChar = ' ';
			this.txtValorIva.ReadOnly = true;
			this.txtValorIva.Size = new System.Drawing.Size(96, 22);
			this.txtValorIva.TabIndex = 171;
			// 
			// txtFondoG
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFondoG.Appearance = appearance8;
			this.txtFondoG.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.FondosG"));
			this.txtFondoG.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFondoG.FormatString = "#,##0.00";
			this.txtFondoG.Location = new System.Drawing.Point(320, 192);
			this.txtFondoG.Name = "txtFondoG";
			this.txtFondoG.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFondoG.PromptChar = ' ';
			this.txtFondoG.ReadOnly = true;
			this.txtFondoG.Size = new System.Drawing.Size(104, 22);
			this.txtFondoG.TabIndex = 170;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(7, 144);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(53, 16);
			this.label8.TabIndex = 169;
			this.label8.Text = "12% IVA:";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(216, 200);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(90, 16);
			this.label9.TabIndex = 168;
			this.label9.Text = "Fondo Garantias:";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_P_PresupuestoSolicitudPago.PresupuestoSolicitudPago - Detalle SP";
			this.ultraGrid1.DataSource = this.cdsSolicitudPago;
			appearance9.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance9;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Capitulo";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 56;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 52;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 63;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 88;
			ultraGridColumn9.Header.Caption = "Valor";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 79;
			ultraGridColumn10.Header.Caption = "Descripcion";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 110;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 352);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(434, 88);
			this.ultraGrid1.TabIndex = 172;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(8, 328);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(196, 16);
			this.label10.TabIndex = 173;
			this.label10.Text = "Instrucciones de Cargo al Presupuesto";
			// 
			// cmbCapitulo
			// 
			this.cmbCapitulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCapitulo.DisplayMember = "";
			this.cmbCapitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCapitulo.Location = new System.Drawing.Point(40, 368);
			this.cmbCapitulo.Name = "cmbCapitulo";
			this.cmbCapitulo.Size = new System.Drawing.Size(175, 48);
			this.cmbCapitulo.TabIndex = 174;
			this.cmbCapitulo.ValueMember = "";
			this.cmbCapitulo.Visible = false;
			// 
			// cdsPago
			// 
			this.cdsPago.BindingContextCtrl = this;
			this.cdsPago.DataLibrary = "LibFacturacion";
			this.cdsPago.DataLibraryUrl = "";
			this.cdsPago.DataSetDef = "dsPresupuestoC";
			this.cdsPago.EnforceConstraints = false;
			this.cdsPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPago.Name = "cdsPago";
			this.cdsPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPago.SchemaDef = null;
			this.cdsPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsPago_BeforeFill);
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(8, 256);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(68, 15);
			this.ultraLabel6.TabIndex = 177;
			this.ultraLabel6.Text = "Cuenta Corr:";
			// 
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.idCuentaCorriente"));
			this.cmbCuentaCorr.DisplayLayout.AutoFitColumns = true;
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(88, 256);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(124, 21);
			this.cmbCuentaCorr.TabIndex = 176;
			this.cmbCuentaCorr.ValueMember = "idCuentaCorriente";
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.Location = new System.Drawing.Point(224, 256);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(72, 18);
			this.ultraLabel4.TabIndex = 179;
			this.ultraLabel4.Text = "Forma Pago:";
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.idFormaPago"));
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(304, 256);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(120, 21);
			this.cmbFormaPago.TabIndex = 178;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.idCliente"));
			this.cmbCliente.DisplayLayout.AutoFitColumns = true;
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(107, 32);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.ReadOnly = true;
			this.cmbCliente.Size = new System.Drawing.Size(317, 21);
			this.cmbCliente.TabIndex = 180;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_P_PresupuestoSolicitudPago";
			this.barraDato1.DataNombreId = "idSolicitudPago";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsSolicitudPago;
			this.barraDato1.DataTabla = "P_PresupuestoSolicitudPago";
			this.barraDato1.DataTablaHija = "P_PresupuestoDetSolicitudPago";
			this.barraDato1.Location = new System.Drawing.Point(8, 448);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(303, 25);
			this.barraDato1.TabIndex = 181;
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
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// ultraTextEditor1
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance15;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.Facturas"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(107, 56);
			this.ultraTextEditor1.Multiline = true;
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.ReadOnly = true;
			this.ultraTextEditor1.Size = new System.Drawing.Size(317, 28);
			this.ultraTextEditor1.TabIndex = 182;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(8, 296);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(37, 16);
			this.label11.TabIndex = 183;
			this.label11.Text = "Notas:";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// txtIdSolicitud
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdSolicitud.Appearance = appearance16;
			this.txtIdSolicitud.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.idSolicitudPago"));
			this.txtIdSolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSolicitud.Location = new System.Drawing.Point(112, 8);
			this.txtIdSolicitud.Name = "txtIdSolicitud";
			this.txtIdSolicitud.Size = new System.Drawing.Size(56, 22);
			this.txtIdSolicitud.TabIndex = 184;
			// 
			// txtRetIva
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetIva.Appearance = appearance17;
			this.txtRetIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.RetIva"));
			this.txtRetIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetIva.FormatString = "#,##0.00";
			this.txtRetIva.Location = new System.Drawing.Point(104, 192);
			this.txtRetIva.Name = "txtRetIva";
			this.txtRetIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetIva.PromptChar = ' ';
			this.txtRetIva.ReadOnly = true;
			this.txtRetIva.Size = new System.Drawing.Size(96, 22);
			this.txtRetIva.TabIndex = 186;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 192);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(40, 16);
			this.label12.TabIndex = 185;
			this.label12.Text = "Ret Iva";
			// 
			// txtRetRenta
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetRenta.Appearance = appearance18;
			this.txtRetRenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.RetFuente"));
			this.txtRetRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetRenta.FormatString = "#,##0.00";
			this.txtRetRenta.Location = new System.Drawing.Point(104, 216);
			this.txtRetRenta.Name = "txtRetRenta";
			this.txtRetRenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetRenta.PromptChar = ' ';
			this.txtRetRenta.ReadOnly = true;
			this.txtRetRenta.Size = new System.Drawing.Size(96, 22);
			this.txtRetRenta.TabIndex = 188;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(8, 216);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(54, 16);
			this.label13.TabIndex = 187;
			this.label13.Text = "Ret Renta";
			// 
			// txtValorPagar
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorPagar.Appearance = appearance19;
			this.txtValorPagar.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.TotalPagar"));
			this.txtValorPagar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorPagar.FormatString = "#,##0.00";
			this.txtValorPagar.Location = new System.Drawing.Point(320, 216);
			this.txtValorPagar.Name = "txtValorPagar";
			this.txtValorPagar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorPagar.PromptChar = ' ';
			this.txtValorPagar.ReadOnly = true;
			this.txtValorPagar.Size = new System.Drawing.Size(104, 22);
			this.txtValorPagar.TabIndex = 190;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(216, 224);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(71, 16);
			this.label14.TabIndex = 189;
			this.label14.Text = "Valor a Pagar";
			// 
			// ultraNumericEditor1
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance20;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.SubExcento"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(104, 120);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.ReadOnly = true;
			this.ultraNumericEditor1.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor1.TabIndex = 192;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(8, 120);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(98, 16);
			this.label15.TabIndex = 191;
			this.label15.Text = "Sub Total IVA 0%:";
			// 
			// txtFacturado
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFacturado.Appearance = appearance21;
			this.txtFacturado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.TotalFactura"));
			this.txtFacturado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFacturado.FormatString = "#,##0.00";
			this.txtFacturado.Location = new System.Drawing.Point(104, 168);
			this.txtFacturado.Name = "txtFacturado";
			this.txtFacturado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFacturado.PromptChar = ' ';
			this.txtFacturado.ReadOnly = true;
			this.txtFacturado.Size = new System.Drawing.Size(96, 22);
			this.txtFacturado.TabIndex = 193;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(8, 168);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(82, 16);
			this.label16.TabIndex = 194;
			this.label16.Text = "Total Facturado";
			// 
			// ultraNumericEditor2
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor2.Appearance = appearance22;
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolicitudPago, "_P_PresupuestoSolicitudPago.SubIva"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(104, 96);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.ReadOnly = true;
			this.ultraNumericEditor2.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor2.TabIndex = 195;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(8, 96);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(77, 16);
			this.label17.TabIndex = 196;
			this.label17.Text = "Sub Total IVA:";
			// 
			// PresupuestoOrdenTrabajo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(440, 478);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.txtFacturado);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtValorPagar);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtRetIva);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.cmbCuentaCorr);
			this.Controls.Add(this.cmbCapitulo);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtValorIva);
			this.Controls.Add(this.txtFondoG);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.txtAnticipos);
			this.Controls.Add(this.txtMultas);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtValorPagos);
			this.Controls.Add(this.txtPagoFecha);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbPago);
			this.Controls.Add(this.txtIdSolicitud);
			this.Controls.Add(this.txtRetRenta);
			this.Controls.Add(this.label13);
			this.Name = "PresupuestoOrdenTrabajo";
			this.Text = "Solicitud de Pago";
			this.Load += new System.EventHandler(this.PresupuestoOrdenTrabajo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSolicitudPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagoFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMultas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCapitulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorPagar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacturado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void PresupuestoOrdenTrabajo_Load(object sender, System.EventArgs e)
		{
			txtIdSolicitud.Width=0;
			txtIdSolicitud.Height=0;
			#region carga combos
			string stSelect ="Select idCapitulo,Capitulo from P_PresupuestoCapitulo";
			cmbCapitulo.DataSource = Funcion.dvProcedimiento(cdsSolicitudPago,stSelect);
			cmbCapitulo.Rows.Band.Columns["idCapitulo"].Hidden = true;
			cmbCapitulo.DisplayMember = "Capitulo";
			cmbCapitulo.ValueMember = "idCapitulo";

			stSelect ="Exec P_PresupuestoCargaPago";
			cmbPago.DataSource = Funcion.dvProcedimiento(cdsPago,stSelect);
			cmbPago.Rows.Band.Columns["idPagoLote"].Hidden = true;
			cmbPago.DisplayMember = "Codigo";
			cmbPago.ValueMember = "idPagoLote";

			stSelect = "select idFormaPago,FormaPago from compraformapago where Borrar = 0 order by FormaPago";
			cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsPago,stSelect);
			cmbFormaPago.Rows.Band.Columns["idFormaPago"].Hidden = true;
			cmbFormaPago.DisplayMember = "FormaPago";
			cmbFormaPago.ValueMember = "idFormaPago";

			stSelect = "select idCuentaCorriente,NumCuenta from CuentaCorriente where Borrar = 0";
			cmbCuentaCorr.DataSource = Funcion.dvProcedimiento(cdsPago,stSelect);
			cmbCuentaCorr.Rows.Band.Columns["idCuentaCorriente"].Hidden = true;
			cmbCuentaCorr.DisplayMember = "NumCuenta";
			cmbCuentaCorr.ValueMember = "idCuentaCorriente";

			stSelect = "select idCliente,Nombre from Cliente where Proveedor = 1 or Ambos = 1";
			cmbCliente.DataSource = Funcion.dvProcedimiento(cdsPago,stSelect);
			cmbCliente.Rows.Band.Columns["idCliente"].Hidden = true;
			cmbCliente.DisplayMember = "Nombre";
			cmbCliente.ValueMember = "idCliente";
			#endregion carga combos
			barraDato1.HabilitaControles(false);
			barraDato1.PosUltima();
			
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			#region Carga detCapitulo 
			if (Funcion.bEjecutaSQL(cdsSolicitudPago, "Exec SeteoGExiste 'FTRPCO'")) 				
			{	
				if (e.Row.Cells["idCapitulo"].Value != DBNull.Value || e.Row.Cells["idCapitulo"].Value.ToString().Length > 0)
				{
					string stSelect = "Select  count(idCapitulo) from P_PresupuestoCapitulo where idCapitulo = "+(int )e.Row.Cells["idCapitulo"].Value;
					if(Funcion.iEscalarSQL(cdsPago,stSelect)>0)
					{
						string stSelect1 = "Select NotasCapitulo from P_PresupuestoCapitulo where idCapitulo = "+(int )e.Row.Cells["idCapitulo"].Value;
						e.Row.Cells["DescripcionCapitulo"].Value = Funcion.sEscalarSQL(cdsPago,stSelect1);
					}
				}

			}
			#endregion Carga detCapitulo 
		}

		private void cdsSolicitudPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSolicitudPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		private C1.Data.C1DataRow drSolicitudP;

		private void cdsSolicitudPago_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_P_PresupuestoSolicitudPago")
			{
				drSolicitudP = e.Row;				
			}
		}

	
		#region Control BarraDato
		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
		  barraDato1.CrearRegistro();
			errorProvider.SetError(cmbPago,"");
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			if (controlPago == 1)controlPago=0;
			barraDato1.DeshacerRegistro();
			errorProvider.SetError(cmbPago,"");

		}

		int controlPago = 0;
		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
			errorProvider.SetError(cmbPago,"");
			controlPago = 1;
			cmbPago.Enabled = false;
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbPago,"");

			if (drSolicitudP == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			 

			if (DialogResult.No == MessageBox.Show("Se borrara de forma permanente este registro....", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			barraDato1.BorraRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbPago,"");

      if (controlPago == 1) 
			{
					double dValor1;
					dValor1 =0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						if (dr.Cells["Valor1"].Value != DBNull.Value)
							dValor1 += Math.Round( (double) dr.Cells["Valor1"].Value,2);
					}

					if ( Math.Abs(  Math.Round(double.Parse ( txtFacturado.Value.ToString()),2)-dValor1)>0.02)
					{
						MessageBox.Show("Total de la grilla:"+ dValor1.ToString()   +" debe ser igual al total Facturado. No se puede grabar esta edicion");  
						return;
					}  
			}

			if (controlPago == 1)   
			{
				try
				{
					barraDato1.GrabaRegistro();
				}
				catch(Exception ex){MessageBox.Show(""+ex.Message);return;}
					controlPago=0;
					return;
				}

			if (cmbPago.Value == null || cmbPago.Value == System.DBNull.Value || cmbPago.Value.ToString().Length < 0) 
			{
				errorProvider.SetError(cmbPago, "Seleccione Pago para crear solicitud");
				return;
			}
			string stSlect = "Select count(*) from P_PresupuestoSolicitudPago where idPago = "+cmbPago.Value.ToString();
			if(Funcion.iEscalarSQL(cdsPago,stSlect)>0)
			{
				errorProvider.SetError(cmbPago,"Pago ya Generado Solicitud...");
				return;
			}
			errorProvider.SetError(cmbPago,"");
			if (DialogResult.No == MessageBox.Show("Creará una nueva solicitud de pago....", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) 
			{
				errorProvider.SetError(cmbPago,"");
				barraDato1.DeshacerRegistro();
				return;
			}
			barraDato1.GrabaRegistro();

			cargaSolicitud();

			barraDato1.ProximoId(5);
//			barraDato1.EditarRegistro();
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			if (txtIdSolicitud.Value == System.DBNull.Value || txtIdSolicitud.Value == null || int.Parse(txtIdSolicitud.Value.ToString()) < 0 )return;

			this.Cursor = Cursors.WaitCursor;
			string stFiltro = "{P_PresupuestoSolicitudPago.idSolicitudPago} = " + txtIdSolicitud.Value;
			string stReporte = "P_SolicitudPago.rpt";			
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			
			this.Cursor = Cursors.Default;
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}
		#endregion Control BarraDato
		private void cargaSolicitud()
		{
			
			if (cmbPago.Value == null || cmbPago.Value == System.DBNull.Value || cmbPago.Value.ToString().Length < 0) 
			{
				errorProvider.SetError(cmbPago, "Seleccione Pago para crear solicitud");
				return;
			}
			Funcion.EjecutaSQL(cdsPago,"Exec P_PresupuestoCrearSolicitudPago "+ txtIdSolicitud.Value);
		}

		private void cmbPago_Validated(object sender, System.EventArgs e)
		{
//			if (barraDato1.bEditar)
//				cargaSolicitud();
		}
		private void cdsPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
int idSolicitud=0;
		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (P_BuscarOrdenPago miBusqueda = new P_BuscarOrdenPago())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{					
					idSolicitud = miBusqueda.idSolicitudPago;
					barraDato1.IdRegistro = idSolicitud;
					barraDato1.ProximoId(5);
				}
			}
		}
	
	}
}

