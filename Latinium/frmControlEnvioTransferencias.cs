using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmControlEnvioTransferencias.
	/// </summary>
	public class frmControlEnvioTransferencias : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblVendedor;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btGR;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbResponsable;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int IBodega = 0;
		int IdTipoFactura = 0;
		string SCompras = "";
		public frmControlEnvioTransferencias(int idTipoFactura, int iBodega, string sCompras)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			IBodega = iBodega;
			IdTipoFactura = idTipoFactura;
			SCompras = sCompras;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDespacho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraTransf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDespacho");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraTransf");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbResponsable1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btGR = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.chkEstado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			this.SuspendLayout();
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(712, 384);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(67, 16);
			this.lblVendedor.TabIndex = 214;
			this.lblVendedor.Text = "Responsable";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblVendedor.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 78;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 89;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 97;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Número";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 126;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 105;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 151;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 118;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 182;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 202;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Responsable";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 212;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 30;
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
																										 ultraGridColumn12});
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.ActiveRowAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1024, 368);
			this.ultraGrid1.TabIndex = 215;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn12.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
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
			// cmbResponsable1
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable1.Appearance = appearance8;
			this.cmbResponsable1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable1.DisplayMember = "Nombre";
			this.cmbResponsable1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable1.Location = new System.Drawing.Point(792, 384);
			this.cmbResponsable1.Name = "cmbResponsable1";
			this.cmbResponsable1.Size = new System.Drawing.Size(64, 21);
			this.cmbResponsable1.TabIndex = 216;
			this.cmbResponsable1.ValueMember = "idPersonal";
			this.cmbResponsable1.Visible = false;
			this.cmbResponsable1.ValueChanged += new System.EventHandler(this.cmbResponsable_ValueChanged);
			// 
			// btGR
			// 
			appearance9.Image = 14;
			this.btGR.Appearance = appearance9;
			this.btGR.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGR.Location = new System.Drawing.Point(864, 384);
			this.btGR.Name = "btGR";
			this.btGR.Size = new System.Drawing.Size(104, 25);
			this.btGR.TabIndex = 217;
			this.btGR.Text = "Guía de Remisión";
			this.btGR.Visible = false;
			this.btGR.Click += new System.EventHandler(this.btGR_Click);
			// 
			// cmbTipoFactura
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance10;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(648, 384);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(56, 21);
			this.cmbTipoFactura.TabIndex = 220;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			this.cmbTipoFactura.Visible = false;
			this.cmbTipoFactura.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbTipoFactura_RowSelected);
			this.cmbTipoFactura.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoFactura_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(616, 384);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(26, 16);
			this.label1.TabIndex = 219;
			this.label1.Text = "Tipo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Visible = false;
			// 
			// cmbResponsable
			// 
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(792, 128);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(200, 112);
			this.cmbResponsable.TabIndex = 221;
			this.cmbResponsable.ValueMember = "idPersonal";
			this.cmbResponsable.Visible = false;
			// 
			// chkEstado
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEstado.Appearance = appearance11;
			this.chkEstado.BackColor = System.Drawing.Color.Transparent;
			this.chkEstado.Location = new System.Drawing.Point(976, 384);
			this.chkEstado.Name = "chkEstado";
			this.chkEstado.Size = new System.Drawing.Size(56, 22);
			this.chkEstado.TabIndex = 222;
			this.chkEstado.Text = "Estado";
			this.chkEstado.Visible = false;
			this.chkEstado.CheckedChanged += new System.EventHandler(this.chkEstado_CheckedChanged);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance12.Image = 14;
			this.btAceptar.Appearance = appearance12;
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 384);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(104, 25);
			this.btAceptar.TabIndex = 223;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance13.Image = 14;
			this.ultraButton1.Appearance = appearance13;
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Location = new System.Drawing.Point(120, 384);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(104, 25);
			this.ultraButton1.TabIndex = 224;
			this.ultraButton1.Text = "&Cancelar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// frmControlEnvioTransferencias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1040, 414);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.chkEstado);
			this.Controls.Add(this.cmbResponsable);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btGR);
			this.Controls.Add(this.cmbResponsable1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblVendedor);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmControlEnvioTransferencias";
			this.ShowInTaskbar = false;
			this.Text = "Despacho de Bodega";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmControlEnvioTransferencias_KeyDown);
			this.Load += new System.EventHandler(this.frmControlEnvioTransferencias_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void CargaGrid()
		{
			
			string sSQL = string.Format("Exec ListaTransferenciasProcesadas {0}, {1}, {2}, '{3}'", (bool)this.chkEstado.Checked, IdTipoFactura, IBodega, SCompras);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void frmControlEnvioTransferencias_Load(object sender, System.EventArgs e)
		{
			string sBodega = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Nombre From Bodega Where Bodega = {0}", IBodega));

			this.Text = this.Text + " " + sBodega;

			this.cmbResponsable.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where idCargo = 16 Order By Nombre Asc");
			this.cmbResponsable.DisplayLayout.Bands[0].Columns["idPersonal"].Hidden = true;
			this.cmbResponsable.DisplayLayout.Bands[0].Columns["Nombre"].Width = 328;

			CargaGrid();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cmbResponsable_ValueChanged(object sender, System.EventArgs e)
		{
			CargaGrid();
		}

		private void btGR_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (this.cmbResponsable1.Text.ToString() == "")
			{
				MessageBox.Show("Seleccione el responsable.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbResponsable1.Select();
				return;
			}
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Estado"].Value) iCont++;
			}
			if (iCont == 0)
			{
				string sMensaje = "";
				if ((int)this.cmbTipoFactura.Value == 42) sMensaje = "Seleccione la(s) Transferencia(s).";
				else sMensaje = "Seleccione la(s) Factura(s).";

				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.ultraGrid1.Select();
				return;
			}
			#endregion Validación

			#region Destino
			string sSQLDestino = "";
			string sDestino = "";
			string sComparaDestino = "";

			#region Ventas
			if ((int)this.cmbTipoFactura.Value == 1)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if ((bool)dr.Cells["Estado"].Value)
					{
						int iDestino = (int)this.ultraGrid1.ActiveRow.Cells["Bodega1"].Value;
						sSQLDestino = string.Format("Select  From Bodega Where Bodega = {0}", iDestino);
						sComparaDestino = Funcion.sEscalarSQL(cdsSeteoF, sSQLDestino);

						if (sDestino.ToString().Length == 0) sDestino = sComparaDestino;
						else 
						{
							sComparaDestino = Funcion.sEscalarSQL(cdsSeteoF, sSQLDestino);
							int iResComp = sDestino.IndexOf(sComparaDestino);

							if (iResComp == 0) sDestino = " - " + sComparaDestino;
						}
					}
				}
			}			
			#endregion Ventas

			#region Transferencias
			if ((int)this.cmbTipoFactura.Value == 42)
			{
        foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if ((bool)dr.Cells["Estado"].Value)
					{
						int iDestino = (int)this.ultraGrid1.ActiveRow.Cells["Bodega1"].Value;
						sSQLDestino = string.Format("Select Nombre From Bodega Where Bodega = {0}", iDestino);
						sComparaDestino = Funcion.sEscalarSQL(cdsSeteoF, sSQLDestino);

						if (sDestino.ToString().Length == 0) sDestino = sComparaDestino;
						else 
						{							
							int iResComp = sDestino.IndexOf(sComparaDestino);

							if (iResComp == -1) sDestino = " - " + sComparaDestino;
						}

					}
				}
			}			
			#endregion Transferencias
			#endregion Destino
			
			
		}

		private void cmbTipoFactura_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipoFactura_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
			CargaGrid();
		}

		private void chkEstado_CheckedChanged(object sender, System.EventArgs e)
		{
			CargaGrid();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		public int iContRows = 0;
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Estado"].Value) iCont++;
			}

			if (iCont == 0)
			{
				MessageBox.Show("Seleccione la(s) transferencia(s)", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.ultraGrid1.Select();
				return;
			}

			#region Valida 20 filas en el detalle
			int iContFilas = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Estado"].Value)
				{			
					int idCompraTransf = (int)dr.Cells["idCompra"].Value;
					int iFilas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetCompra Where idCompra = {0}", idCompraTransf));
					iContFilas = iContFilas + iFilas;
				}
			}

			if (iContFilas > 20)
			{
				MessageBox.Show("Las transferencias seleccionadas sobrepasan de 20 items.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.ultraGrid1.Select();
				return;
			}
			#endregion Valida 20 filas en el detalle

			this.DialogResult = DialogResult.OK;
			Close();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
			Close();
		}

		private void frmControlEnvioTransferencias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}
	}
}

