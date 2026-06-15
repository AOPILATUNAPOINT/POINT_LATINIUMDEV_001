using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPagosSinFactura.
	/// </summary>
	public class frmPagosSinFactura : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbConceptosPagos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbConceptoPagos;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnAnular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBeneficiario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;

		#region Variables

		bool bNuevo = false;
		bool bEdicion = false;
		int IdPagosSinFactura = 0; 
		int idBloqueaTransacciones = 0;
		int idBloqueaTransacciones1 = 0;
		private Acceso miAcceso;
		private bool bActualiza = true;
		bool bAnulado = false;
		int idTipoOrigen = 0;

		#endregion Variables
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPagosSinFactura;
		private System.Windows.Forms.Label lblComprobante;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;


		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPagosSinFactura()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmPagosSinFactura(int idPagosSinFactura)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			IdPagosSinFactura = idPagosSinFactura;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConcepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConcepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConcepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPagosSinFactura));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label3 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbConceptosPagos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbConceptoPagos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtBeneficiario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidPagosSinFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConceptosPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConceptoPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPagosSinFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(101, 16);
			this.label3.TabIndex = 765;
			this.label3.Text = "Concepto de Pago:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// cmbConceptosPagos
			// 
			this.cmbConceptosPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbConceptosPagos.DataSource = this.ultraDataSource1;
			this.cmbConceptosPagos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 20;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 81;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 22;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.SummaryFooterCaption = "Grand Summaries";
			this.cmbConceptosPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbConceptosPagos.DisplayMember = "Concepto";
			this.cmbConceptosPagos.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbConceptosPagos.Location = new System.Drawing.Point(520, 192);
			this.cmbConceptosPagos.Name = "cmbConceptosPagos";
			this.cmbConceptosPagos.Size = new System.Drawing.Size(100, 56);
			this.cmbConceptosPagos.TabIndex = 976;
			this.cmbConceptosPagos.ValueMember = "idConcepto";
			this.cmbConceptosPagos.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cmbConceptoPagos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbConceptoPagos.Appearance = appearance1;
			this.cmbConceptoPagos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbConceptoPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbConceptoPagos.DataSource = this.ultraDataSource1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 344;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbConceptoPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbConceptoPagos.DisplayMember = "Concepto";
			this.cmbConceptoPagos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbConceptoPagos.Enabled = false;
			this.cmbConceptoPagos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbConceptoPagos.Location = new System.Drawing.Point(120, 8);
			this.cmbConceptoPagos.Name = "cmbConceptoPagos";
			this.cmbConceptoPagos.Size = new System.Drawing.Size(340, 21);
			this.cmbConceptoPagos.TabIndex = 986;
			this.cmbConceptoPagos.ValueMember = "idConcepto";
			this.cmbConceptoPagos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbConceptoPagos_KeyDown);
			this.cmbConceptoPagos.Validating += new System.ComponentModel.CancelEventHandler(this.cmbConceptoPagos_Validating);
			this.cmbConceptoPagos.Validated += new System.EventHandler(this.cmbConceptoPagos_Validated);
			this.cmbConceptoPagos.ValueChanged += new System.EventHandler(this.cmbConceptoPagos_ValueChanged);
			this.cmbConceptoPagos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbConceptoPagos_InitializeLayout);
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsReturn = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance2;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(8, 128);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(616, 56);
			this.txtDescripcion.TabIndex = 987;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 104);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(148, 16);
			this.label1.TabIndex = 988;
			this.label1.Text = "Concepto/Descripción/Notas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 192);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(67, 16);
			this.label4.TabIndex = 990;
			this.label4.Text = "Beneficiario:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBeneficiario
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBeneficiario.Appearance = appearance3;
			this.txtBeneficiario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBeneficiario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBeneficiario.Enabled = false;
			this.txtBeneficiario.Location = new System.Drawing.Point(80, 192);
			this.txtBeneficiario.Multiline = true;
			this.txtBeneficiario.Name = "txtBeneficiario";
			this.txtBeneficiario.Size = new System.Drawing.Size(376, 21);
			this.txtBeneficiario.TabIndex = 989;
			this.txtBeneficiario.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBeneficiario_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 224);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 991;
			this.label2.Text = "Valor:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValor
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance4;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(80, 224);
			this.txtValor.MinValue = 0;
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(96, 21);
			this.txtValor.TabIndex = 992;
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(184, 275);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(80, 23);
			this.btnEditar.TabIndex = 996;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(272, 275);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 995;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(448, 275);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 23);
			this.btnCancelar.TabIndex = 997;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(536, 275);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 998;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 275);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 23);
			this.btnNuevo.TabIndex = 993;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(96, 275);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 994;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(360, 274);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(80, 24);
			this.btnAnular.TabIndex = 999;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(464, 72);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 25);
			this.lblEstado.TabIndex = 1000;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(576, 8);
			this.txtEstado.MaxValue = 3600;
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(40, 21);
			this.txtEstado.TabIndex = 1002;
			this.txtEstado.Visible = false;
			// 
			// txtidPagosSinFactura
			// 
			this.txtidPagosSinFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPagosSinFactura.Enabled = false;
			this.txtidPagosSinFactura.Location = new System.Drawing.Point(528, 8);
			this.txtidPagosSinFactura.MaxValue = 3600;
			this.txtidPagosSinFactura.MinValue = 0;
			this.txtidPagosSinFactura.Name = "txtidPagosSinFactura";
			this.txtidPagosSinFactura.PromptChar = ' ';
			this.txtidPagosSinFactura.Size = new System.Drawing.Size(40, 21);
			this.txtidPagosSinFactura.TabIndex = 1003;
			this.txtidPagosSinFactura.Visible = false;
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblComprobante.Location = new System.Drawing.Point(8, 42);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(84, 16);
			this.lblComprobante.TabIndex = 1004;
			this.lblComprobante.Text = "Centro de costo";
			// 
			// cmbProyecto
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance5;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 400;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand3.SummaryFooterCaption = "Grand Summaries";
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbProyecto.Location = new System.Drawing.Point(120, 40);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(340, 21);
			this.cmbProyecto.TabIndex = 1005;
			this.cmbProyecto.ValueMember = "idProyecto";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 74);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(39, 17);
			this.label10.TabIndex = 1006;
			this.label10.Text = "Cuenta";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCuenta
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance6;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 400;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand4.SummaryFooterCaption = "Grand Summaries";
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCuenta.DisplayMember = "Nombre";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCuenta.Location = new System.Drawing.Point(120, 72);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(340, 21);
			this.cmbCuenta.TabIndex = 1007;
			this.cmbCuenta.ValueMember = "idProyecto";
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			// 
			// frmPagosSinFactura
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(632, 304);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.txtidPagosSinFactura);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.txtBeneficiario);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.cmbConceptoPagos);
			this.Controls.Add(this.cmbConceptosPagos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmPagosSinFactura";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Pagos Sin Factura";
			this.Load += new System.EventHandler(this.frmPagosSinFactura_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConceptosPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConceptoPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPagosSinFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}

		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion 		
			if (!Validacion.vbComboVacio(this.cmbConceptoPagos, "Seleccione un Concepto De Pago")) return;
			if (!Validacion.vbComboVacio(this.cmbProyecto, "Seleccione un Centro De Costo")) return;
			if (!Validacion.vbTexto(this.txtDescripcion, 10, 300, "Ingrese la Descripcion")) return;
			if (MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 33)
			{
				if (!Validacion.vbTexto(this.txtBeneficiario, 3, 50, "Ingrese el Beneficiario")) return;
			}
			else
			{
				if (!Validacion.vbTexto(this.txtBeneficiario, 3, 80, "Ingrese el Beneficiario")) return;
			}
			if (miAcceso.EditarCuentaCodigo) 		
			{
				if((int)this.cmbConceptoPagos.Value == 12 )
				{
					if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione una cuenta")) return;
				
				}
			}
			#endregion Validacion

			#region valida Presupuesto
			if (bNuevo)
			{
				if (MenuLatinium.stIdDB == 1)// solo para point 
				{
					int iCentroDeCosto = 0 ;
					
					if (this.cmbProyecto.Value != System.DBNull.Value) iCentroDeCosto = (int)this.cmbProyecto.Value;
					if ( iCentroDeCosto > 0 )
					{
						int idCuentaExi = 0;
						int idCuentasS = 0;

						idCuentasS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaConceptosPagosSinFacttura] {0}",
							(int)this.cmbConceptoPagos.Value));

						idCuentaExi = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCuentaPresupuestoAsiento] {0}, {1}, '{2}'",
							idCuentasS, iCentroDeCosto, DateTime.Today.ToString("yyyyMMdd")));	

						if (idCuentaExi > 0 )
						{
							int iProyectoS = 0;
							iProyectoS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCCPresupuestoAsiento] {0}, '{1}'",
								iCentroDeCosto,DateTime.Today.ToString("yyyyMMdd") ));

							if ( iProyectoS > 0 )
							{
								#region valida valores
								decimal Presupuesto = 0;
								decimal Utilizado = 0;
								decimal Disponible = 0;
								string Cuenta = "";
								if ( Convert.ToDecimal ( this.txtValor.Value)  > 0 )
								{
									string sSQL = string.Format("Exec [ValidaCCPresupuestoAsientoValores] '{0}', {1}, {2}", 
										DateTime.Today.ToString("yyyyMMdd"), idCuentasS, (int) iCentroDeCosto );
									SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
									dra.Read();
									if (dra.HasRows)
									{
										if (dra.GetValue(0) != System.DBNull.Value)Presupuesto = dra.GetDecimal (0);
										if (dra.GetValue(1) != System.DBNull.Value)Utilizado = dra.GetDecimal (1);
										if (dra.GetValue(2) != System.DBNull.Value)Disponible = dra.GetDecimal (2);
										if (dra.GetValue(3) != System.DBNull.Value)Cuenta = dra.GetString (3);
											
									}
									dra.Close();
								}
								if ( Convert.ToDecimal ( this.txtValor.Value)  > Disponible )
								{
									MessageBox.Show(string.Format("Ha superado el presupuesto de '{0}' \n\n Presupuesto : {1}  \n\n Utilizado : {2} \n\n Solicitado : {3}", Cuenta, Presupuesto, Utilizado, Convert.ToDecimal ( this.txtValor.Value) ), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);											
									return;
								}
								#endregion valida valores
							}           
						}
					}
				}
			}
			#endregion valida Presupuesto
			#region Guardar
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Graba Maestro

					int iCuenta =0;
					int bodega =0;

					int IdConceptoPago = 0;
			
					if (this.cmbConceptoPagos.ActiveRow != null) IdConceptoPago  = (int)this.cmbConceptoPagos.Value;

					if (miAcceso.EditarCuentaCodigo) 		
					{	
						if((int)this.cmbConceptoPagos.Value == 12 )
					{
						iCuenta=(int)this.cmbCuenta.Value;
					}
					else 
					{
								 iCuenta=0;
							 }
					}
					string  sSQL1 = string.Format("Exec GrabaPagosSinFactura {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7},{8}" ,
						(int)this.txtidPagosSinFactura.Value,
						22,	IdConceptoPago,
						this.txtDescripcion.Text,
						this.txtBeneficiario.Text,
						this.txtValor.Value, 
						(int)txtEstado.Value,
						(int)this.cmbProyecto.Value,
						iCuenta);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidPagosSinFactura.Value = (int)oCmdActualiza.ExecuteScalar();
					Estado((int)this.txtEstado.Value);
					#endregion Graba Maestro

					#region Autorización de Pagos					
					string sSQL = string.Format("Exec CreaAutorizacionesDePago 22, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'PAGO SIN FACTURA N. DOCUMENTO {4}', 0", 
						(int)this.txtidPagosSinFactura.Value, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"), Convert.ToDecimal(this.txtValor.Value), this.cmbConceptoPagos.Text.ToString());
					oCmdActualiza.CommandText = sSQL;
					oCmdActualiza.ExecuteNonQuery();	
					#endregion Autorización de Pagos

					#region Bloqueo
					if (!bNuevo)
					{
						string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtidPagosSinFactura.Value);
						oCmdActualiza.CommandText = sSQLBloqueo;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Bloqueo
					
					oTransaction.Commit();

					oConexion.Close();

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					#region Controles
					this.btnGuardar.Enabled = false;
					this.cmbConceptoPagos.Enabled = false;
					this.cmbCuenta.Enabled = false;
					this.cmbProyecto.Enabled = false;
					this.btnNuevo.Enabled = true;
					this.btnConsultar.Enabled = true;
					this.btnEditar.Enabled = true;
					this.txtDescripcion.Enabled = false;
					this.txtBeneficiario.Enabled = false;
					this.txtValor.Enabled = false;
					bNuevo = false;
					bEdicion = false;

					#endregion Controles
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
			#endregion Guardar
		}

		private void frmPagosSinFactura_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0445");

			if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;

			this.cmbConceptoPagos.DataSource = FuncionesProcedimientos.dtGeneral("Select idConcepto, Concepto, c.Descripcion As Cuenta From ConceptosPagosSinFacttura p Inner Join Cuenta c On c.idCuenta = p.idCuenta  where  activo=1");
			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

		  Limpiar();

			using (frmBuscaPagosSinFactura Busqueda = new frmBuscaPagosSinFactura((int)this.txtidPagosSinFactura.Value))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					ConsultaRegistro((int)Busqueda.ultraGrid1.ActiveRow.Cells["idPagosSinFactura"].Value);
				}
			}
		}

		private void ConsultaRegistro (int IdPagosSinFactura)
		{
			#region botones 
			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.cmbConceptoPagos.Enabled = false;
			
			this.txtDescripcion.Enabled = false;
			this.txtBeneficiario.Enabled = false;
			this.txtValor.Enabled = false;
			this.cmbCuenta.Enabled = false;

			#endregion botones

			this.txtidPagosSinFactura.Value = IdPagosSinFactura;

			string sSQL = string.Format("Exec ConsultaIndividualPagosSinFactura {0}", (int)this.txtidPagosSinFactura.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidPagosSinFactura.Value = dr.GetInt32(0);
				this.cmbConceptoPagos.Value = dr.GetInt32(1);
				this.txtDescripcion.Text = dr.GetString(2);
				this.txtBeneficiario.Text = dr.GetString(3);
				this.txtValor.Value = dr.GetDecimal(4);
				this.txtEstado.Value = dr.GetInt32(5);
				this.cmbProyecto.Value = dr.GetInt32(6);
				this.cmbCuenta.Value = dr.GetInt32(7);
			}
			dr.Close();

			Estado ((int) this.txtEstado.Value);
			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.cmbConceptoPagos.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.cmbProyecto.Enabled = false;
			this.txtEstado.Value = 0;
			this.txtDescripcion.Enabled = false;
			this.txtBeneficiario.Enabled = false;
			this.txtValor.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;		
			Limpiar();	
			bAnulado = false;		
			this.btnNuevo.Focus();

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			bNuevo = true;
			bEdicion = true;

			this.cmbConceptoPagos.Enabled = true;
			this.cmbProyecto.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.txtEstado.Value = 1;
			this.txtidPagosSinFactura.Value = 0;
			
			this.txtDescripcion.Enabled = true;
			this.txtBeneficiario.Enabled = true;
			this.txtValor.Enabled = true;
			this.cmbConceptoPagos.Focus ();
		
		}

		private void Estado ( int IEstado)
		{
			if (IEstado == 1)
				this.lblEstado.Text = "PENDIENTE";
			
			if (IEstado == 2)
				this.lblEstado.Text = " AUTORIZADO";
			
			if (IEstado == 3)
				this.lblEstado.Text = "PAGADO";

			if (IEstado == 4)
			{
				this.lblEstado.Text = "ANULADO";
				this.btnEditar.Enabled = false;
				this.btnAnular.Enabled = false;
			}
		}

		private void Limpiar(){
			
			this.txtDescripcion.Text= "";
			this.txtBeneficiario.Text= "";
			this.txtValor.Value = 0;
			this.txtEstado.Value = 0;
			this.lblEstado.Text = "";
      this.txtidPagosSinFactura.Value  = 0;
			this.cmbConceptoPagos.Value = System.DBNull.Value;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.cmbProyecto.Value = System.DBNull.Value;

		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.ConsultaRegistro  ((int)this.txtidPagosSinFactura.Value);

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("No puede editar el documento esta Autorizado su Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede editar el documento esta Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("No puede editar el documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else {
				#region botones
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = true;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = true;
				this.cmbConceptoPagos.Enabled = true;
				//this.cmbCuenta.Enabled = false;
				if (miAcceso.EditarCuentaCodigo) 		
				{
					this.cmbCuenta.Enabled = true;
				}
				this.cmbProyecto.Enabled = true;
				this.txtDescripcion.Enabled = true;
				this.txtBeneficiario.Enabled = true;
				this.txtValor.Enabled = true;

				this.cmbConceptoPagos.Focus();
				#endregion Botones

				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtidPagosSinFactura.Value);
				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);	

				bNuevo = false;
				bEdicion = true;
			}	
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{	
			if ((int) this.txtidPagosSinFactura.Value == 0) return;
			
			this.ConsultaRegistro((int)this.txtidPagosSinFactura.Value);			

			if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("El Documento ya esta Anulado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede Anular el Documento Esta Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}		
	
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("No puede Anular el Documento Esta Autorizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de ANULAR el documento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update PagosSinFactura Set Estado = 4 Where idPagosSinFactura = {0}", (int)this.txtidPagosSinFactura.Value));
		
				this.btnEditar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnGuardar.Enabled = false;

				this.cmbConceptoPagos.Enabled = false;
				this.cmbCuenta.Enabled = false;
				this.cmbProyecto.Enabled = false;
				this.txtDescripcion.Enabled = false;
				this.txtBeneficiario.Enabled = false;
				this.txtValor.Enabled = false;
				this.lblEstado.Text ="ANULADO";
			}
		}

		private void cmbConceptoPagos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDescripcion.Focus();
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtBeneficiario.Focus();
		}

		private void txtBeneficiario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtValor.Focus();
	
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cmbConceptoPagos_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbConceptoPagos.ActiveRow != null)
			{
				this.cmbCuenta.Enabled = false;

				if (miAcceso.EditarCuentaCodigo) 		
				{
					if((int)this.cmbConceptoPagos.Value == 12)
					{
						if (bEdicion) this.cmbCuenta.Enabled = true;
						this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuantasOtrosGastos");
						return;	
					}					
				}
			}
		}

		private void cmbConceptoPagos_Validated(object sender, System.EventArgs e)
		{
			if ((int)this.cmbConceptoPagos.Value != 12)
			{
				this.cmbCuenta.Enabled=false;
				this.cmbCuenta.Value = System.DBNull.Value;
			}
		}

		private void cmbConceptoPagos_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbConceptoPagos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	}
}
