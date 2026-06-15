using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for Cuota_Point.
	/// </summary>
	public class Cuota_Point : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCuotaMin;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDiasMaxiCuotas;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDiasMinCuota;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDiasMaxPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCuotaMini;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValorMinCuota;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCuotaMax;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnInteres;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnValMaxCred;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTasaEfectiva;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVMSolidario;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Cuota_Point()
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
			this.label36 = new System.Windows.Forms.Label();
			this.spnCuotaMin = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsCuotas = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.spnCuotaMax = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.spnDiasMaxiCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.spnDiasMinCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.spnDiasMaxPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.spnCuotaMini = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.spnValorMinCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.spnInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.panel1 = new System.Windows.Forms.Panel();
			this.panel2 = new System.Windows.Forms.Panel();
			this.label10 = new System.Windows.Forms.Label();
			this.spnValMaxCred = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.spnTasaEfectiva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtVMSolidario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.spnCuotaMin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCuotaMax)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasMaxiCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasMinCuota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasMaxPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCuotaMini)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValorMinCuota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnInteres)).BeginInit();
			this.panel1.SuspendLayout();
			this.panel2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.spnValMaxCred)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTasaEfectiva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVMSolidario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(24, 40);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(35, 16);
			this.label36.TabIndex = 74;
			this.label36.Text = "Desde";
			// 
			// spnCuotaMin
			// 
			this.spnCuotaMin.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.MinimoCuotas"));
			this.spnCuotaMin.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCuotaMin.FormatString = "####";
			this.spnCuotaMin.Location = new System.Drawing.Point(136, 40);
			this.spnCuotaMin.MinValue = 0;
			this.spnCuotaMin.Name = "spnCuotaMin";
			this.spnCuotaMin.PromptChar = ' ';
			this.spnCuotaMin.Size = new System.Drawing.Size(88, 22);
			this.spnCuotaMin.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnCuotaMin.TabIndex = 1;
			// 
			// cdsCuotas
			// 
			this.cdsCuotas.BindingContextCtrl = this;
			this.cdsCuotas.DataLibrary = "LibFacturacion";
			this.cdsCuotas.DataLibraryUrl = "";
			this.cdsCuotas.DataSetDef = "dsCuotas";
			this.cdsCuotas.EnforceConstraints = false;
			this.cdsCuotas.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuotas.Name = "cdsCuotas";
			this.cdsCuotas.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCuotas.SchemaDef = null;
			this.cdsCuotas.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuotas_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(135, 16);
			this.label1.TabIndex = 75;
			this.label1.Text = "Numero de cuotas/ meses";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(272, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 77;
			this.label2.Text = "Hasta";
			// 
			// spnCuotaMax
			// 
			this.spnCuotaMax.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.MaximoCuotas"));
			this.spnCuotaMax.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCuotaMax.Location = new System.Drawing.Point(320, 40);
			this.spnCuotaMax.MinValue = 0;
			this.spnCuotaMax.Name = "spnCuotaMax";
			this.spnCuotaMax.PromptChar = ' ';
			this.spnCuotaMax.Size = new System.Drawing.Size(88, 22);
			this.spnCuotaMax.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnCuotaMax.TabIndex = 2;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(115, 16);
			this.label3.TabIndex = 78;
			this.label3.Text = "Fechas de Pago / Dias";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(272, 32);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 16);
			this.label4.TabIndex = 82;
			this.label4.Text = "Hasta";
			// 
			// spnDiasMaxiCuotas
			// 
			this.spnDiasMaxiCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.MaximoDiasCuota"));
			this.spnDiasMaxiCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDiasMaxiCuotas.Location = new System.Drawing.Point(320, 32);
			this.spnDiasMaxiCuotas.MaxValue = 100;
			this.spnDiasMaxiCuotas.MinValue = 1;
			this.spnDiasMaxiCuotas.Name = "spnDiasMaxiCuotas";
			this.spnDiasMaxiCuotas.PromptChar = ' ';
			this.spnDiasMaxiCuotas.Size = new System.Drawing.Size(88, 22);
			this.spnDiasMaxiCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDiasMaxiCuotas.TabIndex = 4;
			this.spnDiasMaxiCuotas.Value = 1;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 16);
			this.label5.TabIndex = 80;
			this.label5.Text = "Desde";
			// 
			// spnDiasMinCuota
			// 
			this.spnDiasMinCuota.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.MinimoDiasCuota"));
			this.spnDiasMinCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDiasMinCuota.Location = new System.Drawing.Point(144, 32);
			this.spnDiasMinCuota.MaxValue = 31;
			this.spnDiasMinCuota.MinValue = 1;
			this.spnDiasMinCuota.Name = "spnDiasMinCuota";
			this.spnDiasMinCuota.PromptChar = ' ';
			this.spnDiasMinCuota.Size = new System.Drawing.Size(88, 22);
			this.spnDiasMinCuota.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDiasMinCuota.TabIndex = 3;
			this.spnDiasMinCuota.Value = 1;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 72);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(106, 16);
			this.label6.TabIndex = 84;
			this.label6.Text = "Dia Maximo de Pago";
			// 
			// spnDiasMaxPago
			// 
			this.spnDiasMaxPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.DiaMaximoPagoCuota"));
			this.spnDiasMaxPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDiasMaxPago.Location = new System.Drawing.Point(144, 64);
			this.spnDiasMaxPago.MaxValue = 31;
			this.spnDiasMaxPago.MinValue = 1;
			this.spnDiasMaxPago.Name = "spnDiasMaxPago";
			this.spnDiasMaxPago.PromptChar = ' ';
			this.spnDiasMaxPago.Size = new System.Drawing.Size(88, 22);
			this.spnDiasMaxPago.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDiasMaxPago.TabIndex = 5;
			this.spnDiasMaxPago.Value = 1;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(88, 304);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(101, 16);
			this.label7.TabIndex = 88;
			this.label7.Text = "Valor Cuota Minima";
			// 
			// spnCuotaMini
			// 
			this.spnCuotaMini.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.CuotaMinima"));
			this.spnCuotaMini.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCuotaMini.FormatString = "#,##0.00";
			this.spnCuotaMini.Location = new System.Drawing.Point(224, 304);
			this.spnCuotaMini.MinValue = 0;
			this.spnCuotaMini.Name = "spnCuotaMini";
			this.spnCuotaMini.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCuotaMini.PromptChar = ' ';
			this.spnCuotaMini.Size = new System.Drawing.Size(152, 22);
			this.spnCuotaMini.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnCuotaMini.TabIndex = 8;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(88, 208);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(124, 16);
			this.label8.TabIndex = 86;
			this.label8.Text = "Valor Minimo de Credito";
			// 
			// spnValorMinCuota
			// 
			this.spnValorMinCuota.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.ValorMinimo"));
			this.spnValorMinCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnValorMinCuota.FormatString = "#,##0.00";
			this.spnValorMinCuota.Location = new System.Drawing.Point(224, 208);
			this.spnValorMinCuota.MinValue = 0;
			this.spnValorMinCuota.Name = "spnValorMinCuota";
			this.spnValorMinCuota.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnValorMinCuota.PromptChar = ' ';
			this.spnValorMinCuota.Size = new System.Drawing.Size(152, 22);
			this.spnValorMinCuota.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnValorMinCuota.TabIndex = 6;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(88, 336);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(40, 16);
			this.label9.TabIndex = 90;
			this.label9.Text = "Interes";
			// 
			// spnInteres
			// 
			this.spnInteres.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.InteresCuota"));
			this.spnInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnInteres.FormatString = "#,##0.00";
			this.spnInteres.Location = new System.Drawing.Point(224, 336);
			this.spnInteres.MaxValue = 100;
			this.spnInteres.MinValue = 0;
			this.spnInteres.Name = "spnInteres";
			this.spnInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnInteres.PromptChar = ' ';
			this.spnInteres.Size = new System.Drawing.Size(152, 22);
			this.spnInteres.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnInteres.TabIndex = 9;
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(120, 408);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(112, 25);
			this.btAceptar.TabIndex = 10;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(248, 408);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(104, 25);
			this.btCancelar.TabIndex = 11;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click_1);
			// 
			// panel1
			// 
			this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel1.Controls.Add(this.label2);
			this.panel1.Controls.Add(this.spnCuotaMax);
			this.panel1.Controls.Add(this.spnCuotaMin);
			this.panel1.Controls.Add(this.label36);
			this.panel1.Controls.Add(this.label1);
			this.panel1.Location = new System.Drawing.Point(8, 8);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(448, 72);
			this.panel1.TabIndex = 0;
			// 
			// panel2
			// 
			this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.panel2.Controls.Add(this.label3);
			this.panel2.Controls.Add(this.label5);
			this.panel2.Controls.Add(this.spnDiasMinCuota);
			this.panel2.Controls.Add(this.label4);
			this.panel2.Controls.Add(this.spnDiasMaxiCuotas);
			this.panel2.Controls.Add(this.label6);
			this.panel2.Controls.Add(this.spnDiasMaxPago);
			this.panel2.Location = new System.Drawing.Point(8, 88);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(448, 96);
			this.panel2.TabIndex = 96;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(88, 240);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(124, 16);
			this.label10.TabIndex = 98;
			this.label10.Text = "Valor Máxmo de Credito";
			// 
			// spnValMaxCred
			// 
			this.spnValMaxCred.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.ValorMaximo"));
			this.spnValMaxCred.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnValMaxCred.FormatString = "#,##0.00";
			this.spnValMaxCred.Location = new System.Drawing.Point(224, 240);
			this.spnValMaxCred.MinValue = 0;
			this.spnValMaxCred.Name = "spnValMaxCred";
			this.spnValMaxCred.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnValMaxCred.PromptChar = ' ';
			this.spnValMaxCred.Size = new System.Drawing.Size(152, 22);
			this.spnValMaxCred.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnValMaxCred.TabIndex = 7;
			this.spnValMaxCred.ValueChanged += new System.EventHandler(this.spnValMaxCred_ValueChanged);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(88, 368);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(102, 16);
			this.label11.TabIndex = 100;
			this.label11.Text = "Tasa Efectiva Anual";
			// 
			// spnTasaEfectiva
			// 
			this.spnTasaEfectiva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.TasaEfectivaAnual"));
			this.spnTasaEfectiva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnTasaEfectiva.FormatString = "#,##0.00";
			this.spnTasaEfectiva.Location = new System.Drawing.Point(224, 368);
			this.spnTasaEfectiva.MaxValue = 100;
			this.spnTasaEfectiva.MinValue = 0;
			this.spnTasaEfectiva.Name = "spnTasaEfectiva";
			this.spnTasaEfectiva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnTasaEfectiva.PromptChar = ' ';
			this.spnTasaEfectiva.Size = new System.Drawing.Size(152, 22);
			this.spnTasaEfectiva.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnTasaEfectiva.TabIndex = 99;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(88, 272);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(115, 16);
			this.label12.TabIndex = 102;
			this.label12.Text = "Valor Máxmo Solidario";
			// 
			// txtVMSolidario
			// 
			this.txtVMSolidario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuotas, "Cuotas.MaximoCuotasUB"));
			this.txtVMSolidario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVMSolidario.FormatString = "#,##0.00";
			this.txtVMSolidario.Location = new System.Drawing.Point(224, 272);
			this.txtVMSolidario.MinValue = 0;
			this.txtVMSolidario.Name = "txtVMSolidario";
			this.txtVMSolidario.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVMSolidario.PromptChar = ' ';
			this.txtVMSolidario.Size = new System.Drawing.Size(152, 22);
			this.txtVMSolidario.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtVMSolidario.TabIndex = 101;
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
			// Cuota_Point
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(464, 440);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtVMSolidario);
			this.Controls.Add(this.spnTasaEfectiva);
			this.Controls.Add(this.spnValMaxCred);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.spnInteres);
			this.Controls.Add(this.spnCuotaMini);
			this.Controls.Add(this.spnValorMinCuota);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Cuota_Point";
			this.ShowInTaskbar = false;
			this.Text = "Parámetros de crédito";
			this.Load += new System.EventHandler(this.Cuota_Point_Load);
			((System.ComponentModel.ISupportInitialize)(this.spnCuotaMin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCuotaMax)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasMaxiCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasMinCuota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasMaxPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCuotaMini)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnValorMinCuota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnInteres)).EndInit();
			this.panel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.spnValMaxCred)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTasaEfectiva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVMSolidario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void Cuota_Point_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0841");

			if (!Funcion.Permiso("1942", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Parametros de Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
				Cursor = Cursors.Default;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{			
			int iMinimoCuotas = (int) spnCuotaMin.Value;
			int iMaximoCuotas = (int) spnCuotaMax.Value;
			int iMinimoDiasCuota = (int) spnDiasMinCuota.Value;
			int iMaximoDiasCuota = (int) spnDiasMaxiCuotas.Value;
			int iDiasMaxPago = (int) spnDiasMaxPago.Value;
			double dValorMinimo = (double) spnValorMinCuota.Value;
			double dValorMaximo = (double) spnValMaxCred.Value;
			double dCuotaMinima = (double) spnCuotaMini.Value;
			double dInteresCuota = (double)spnInteres.Value;

			string stSelect = string.Format("Exec Point_Cuotas {0}, {1},{2},{3},{4},{5},{6},{7},{8},{9}", iMinimoCuotas, iMaximoCuotas,iMinimoDiasCuota,iMaximoDiasCuota,iDiasMaxPago,dValorMinimo,dValorMaximo,dCuotaMinima,dInteresCuota, (double)this.txtVMSolidario.Value);
			string sCuotas= Funcion.sEscalarSQL(cdsCuotas, stSelect, true);

			MessageBox.Show("Datos Grabados correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void cdsCuotas_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuotas.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCancelar_Click_1(object sender, System.EventArgs e)
		{
			Close();
		}

		private void spnValMaxCred_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}
	}
}

