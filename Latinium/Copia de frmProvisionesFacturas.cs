using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmProvisionesFacturas.
	/// </summary>
	public class frmProvisionesFacturas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridProvisiones;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdProvision;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSaldosBancarios;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private System.Windows.Forms.Button btnAutorizar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSumas;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnIrA;
		private System.Windows.Forms.ContextMenu mnuIrA;
		private System.Windows.Forms.MenuItem mnuIrADocumento;
		private System.Windows.Forms.MenuItem mnuIrAEstadoDeCuenta;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.ComboBox cboFecha;
		private System.Windows.Forms.Button btnActualizar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro2;
		private System.Windows.Forms.Button btnCerrar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmProvisionesFacturas()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProvisionesFacturas));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvisionesSaldosBancarios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoAFavor");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProveedores");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoFinal");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PagoProveedores", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PagoProveedores", 4, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoFinal", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoFinal", 5, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoAFavor", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoAFavor", 3, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvisionesSaldosBancarios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoAFavor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PagoProveedores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoFinal");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			this.uGridProvisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtIdProvision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.uGridSaldosBancarios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExcel = new System.Windows.Forms.Button();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.uGridSumas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnIrA = new System.Windows.Forms.Button();
			this.mnuIrA = new System.Windows.Forms.ContextMenu();
			this.mnuIrADocumento = new System.Windows.Forms.MenuItem();
			this.mnuIrAEstadoDeCuenta = new System.Windows.Forms.MenuItem();
			this.lblContador = new System.Windows.Forms.Label();
			this.cboFecha = new System.Windows.Forms.ComboBox();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.optFiltro1 = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optFiltro2 = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnCerrar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridProvisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProvision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSaldosBancarios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSumas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro2)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridProvisiones
			// 
			this.uGridProvisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridProvisiones.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridProvisiones.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridProvisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridProvisiones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridProvisiones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridProvisiones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridProvisiones.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridProvisiones.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridProvisiones.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridProvisiones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridProvisiones.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridProvisiones.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridProvisiones.Font = new System.Drawing.Font("Tahoma", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridProvisiones.Location = new System.Drawing.Point(8, 192);
			this.uGridProvisiones.Name = "uGridProvisiones";
			this.uGridProvisiones.Size = new System.Drawing.Size(1380, 272);
			this.uGridProvisiones.TabIndex = 736;
			this.uGridProvisiones.DoubleClick += new System.EventHandler(this.uGridProvisiones_DoubleClick);
			this.uGridProvisiones.AfterRowsDeleted += new System.EventHandler(this.uGridProvisiones_AfterRowsDeleted);
			this.uGridProvisiones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridProvisiones_KeyPress);
			this.uGridProvisiones.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridProvisiones_BeforeCellUpdate);
			this.uGridProvisiones.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridProvisiones_BeforeRowsDeleted);
			this.uGridProvisiones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridProvisiones_KeyDown);
			this.uGridProvisiones.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridProvisiones_ClickCellButton);
			this.uGridProvisiones.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridProvisiones_AfterCellUpdate);
			this.uGridProvisiones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridProvisiones_InitializeLayout);
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
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Enabled = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(328, 72);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 23);
			this.btnVer.TabIndex = 738;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 75);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 17);
			this.label1.TabIndex = 739;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance7;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Enabled = false;
			this.txtBuscar.Location = new System.Drawing.Point(64, 72);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(250, 22);
			this.txtBuscar.TabIndex = 737;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 176);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1392, 8);
			this.groupBox1.TabIndex = 740;
			this.groupBox1.TabStop = false;
			// 
			// txtIdProvision
			// 
			this.txtIdProvision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdProvision.Enabled = false;
			this.txtIdProvision.Location = new System.Drawing.Point(1304, 8);
			this.txtIdProvision.Name = "txtIdProvision";
			this.txtIdProvision.PromptChar = ' ';
			this.txtIdProvision.Size = new System.Drawing.Size(16, 22);
			this.txtIdProvision.TabIndex = 744;
			this.txtIdProvision.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(200, 7);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 743;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 17);
			this.label2.TabIndex = 742;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGrupo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance8;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.Enabled = false;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(64, 40);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(250, 21);
			this.cmbGrupo.TabIndex = 745;
			this.cmbGrupo.ValueMember = "idGrupoCliente";
			this.cmbGrupo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGrupo_KeyDown);
			this.cmbGrupo.ValueChanged += new System.EventHandler(this.cmbGrupo_ValueChanged);
			this.cmbGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbGrupo_InitializeLayout);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(8, 42);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(35, 17);
			this.label17.TabIndex = 746;
			this.label17.Text = "Grupo";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridSaldosBancarios
			// 
			this.uGridSaldosBancarios.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSaldosBancarios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSaldosBancarios.DataSource = this.ultraDataSource1;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridSaldosBancarios.DisplayLayout.Appearance = appearance9;
			this.uGridSaldosBancarios.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 318;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 212;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance10;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Saldo A Favor";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance11;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Pagos";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 90;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance12;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Saldo Final";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 90;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand3.Header.Caption = "Saldo Bancario";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance13;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance14;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance15;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridSaldosBancarios.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSaldosBancarios.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridSaldosBancarios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSaldosBancarios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridSaldosBancarios.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSaldosBancarios.DisplayLayout.Override.CellAppearance = appearance18;
			this.uGridSaldosBancarios.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 8.5F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSaldosBancarios.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldosBancarios.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldosBancarios.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSaldosBancarios.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridSaldosBancarios.Font = new System.Drawing.Font("Tahoma", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSaldosBancarios.Location = new System.Drawing.Point(884, 8);
			this.uGridSaldosBancarios.Name = "uGridSaldosBancarios";
			this.uGridSaldosBancarios.Size = new System.Drawing.Size(503, 160);
			this.uGridSaldosBancarios.TabIndex = 747;
			this.uGridSaldosBancarios.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSaldosBancarios_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn4.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn5.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(788, 79);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 23);
			this.btnExcel.TabIndex = 748;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// optFiltro
			// 
			this.optFiltro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance23;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 0;
			this.optFiltro.ItemAppearance = appearance24;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Saldo Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Saldo Por Vencer";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.Items.Add(valueListItem3);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(480, 8);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(296, 24);
			this.optFiltro.TabIndex = 749;
			this.optFiltro.Text = "Todos";
			this.optFiltro.ValueChanged += new System.EventHandler(this.optFiltro_ValueChanged);
			// 
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAutorizar.CausesValidation = false;
			this.btnAutorizar.Enabled = false;
			this.btnAutorizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAutorizar.Image")));
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.Location = new System.Drawing.Point(788, 7);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.Size = new System.Drawing.Size(80, 23);
			this.btnAutorizar.TabIndex = 750;
			this.btnAutorizar.Text = "&Autorizar";
			this.btnAutorizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// uGridSumas
			// 
			this.uGridSumas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance25.BackColor = System.Drawing.Color.White;
			this.uGridSumas.DisplayLayout.Appearance = appearance25;
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridSumas.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridSumas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSumas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSumas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.uGridSumas.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSumas.DisplayLayout.Override.CellAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSumas.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.uGridSumas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSumas.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSumas.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.uGridSumas.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSumas.Location = new System.Drawing.Point(8, 104);
			this.uGridSumas.Name = "uGridSumas";
			this.uGridSumas.Size = new System.Drawing.Size(462, 40);
			this.uGridSumas.TabIndex = 751;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(788, 55);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 23);
			this.btnImprimir.TabIndex = 752;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnIrA
			// 
			this.btnIrA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnIrA.Enabled = false;
			this.btnIrA.Image = ((System.Drawing.Image)(resources.GetObject("btnIrA.Image")));
			this.btnIrA.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIrA.Location = new System.Drawing.Point(788, 31);
			this.btnIrA.Name = "btnIrA";
			this.btnIrA.Size = new System.Drawing.Size(80, 23);
			this.btnIrA.TabIndex = 753;
			this.btnIrA.Text = "Ir a";
			this.btnIrA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnIrA.Click += new System.EventHandler(this.btnIrA_Click);
			// 
			// mnuIrA
			// 
			this.mnuIrA.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					 this.mnuIrADocumento,
																																					 this.mnuIrAEstadoDeCuenta});
			// 
			// mnuIrADocumento
			// 
			this.mnuIrADocumento.Index = 0;
			this.mnuIrADocumento.Text = "Documento";
			this.mnuIrADocumento.Click += new System.EventHandler(this.mnuIrADocumento_Click);
			// 
			// mnuIrAEstadoDeCuenta
			// 
			this.mnuIrAEstadoDeCuenta.Index = 1;
			this.mnuIrAEstadoDeCuenta.Text = "Estado De Cuenta";
			this.mnuIrAEstadoDeCuenta.Click += new System.EventHandler(this.mnuIrAEstadoDeCuenta_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 152);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 754;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cboFecha
			// 
			this.cboFecha.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboFecha.Enabled = false;
			this.cboFecha.Location = new System.Drawing.Point(64, 8);
			this.cboFecha.MaxDropDownItems = 15;
			this.cboFecha.Name = "cboFecha";
			this.cboFecha.Size = new System.Drawing.Size(112, 21);
			this.cboFecha.TabIndex = 756;
			this.cboFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cboFecha_KeyDown);
			this.cboFecha.SelectedIndexChanged += new System.EventHandler(this.cboFecha_SelectedIndexChanged);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Enabled = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(788, 103);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(80, 23);
			this.btnActualizar.TabIndex = 757;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// optFiltro1
			// 
			this.optFiltro1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro1.Appearance = appearance31;
			this.optFiltro1.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro1.CheckedIndex = 0;
			this.optFiltro1.Enabled = false;
			this.optFiltro1.ItemAppearance = appearance32;
			this.optFiltro1.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Seleccionado";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "No Seleccionado";
			this.optFiltro1.Items.Add(valueListItem4);
			this.optFiltro1.Items.Add(valueListItem5);
			this.optFiltro1.Items.Add(valueListItem6);
			this.optFiltro1.ItemSpacingVertical = 10;
			this.optFiltro1.Location = new System.Drawing.Point(480, 40);
			this.optFiltro1.Name = "optFiltro1";
			this.optFiltro1.Size = new System.Drawing.Size(296, 24);
			this.optFiltro1.TabIndex = 758;
			this.optFiltro1.Text = "Todos";
			this.optFiltro1.ValueChanged += new System.EventHandler(this.optFiltro1_ValueChanged);
			// 
			// optFiltro2
			// 
			this.optFiltro2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro2.Appearance = appearance33;
			this.optFiltro2.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro2.CheckedIndex = 1;
			this.optFiltro2.Enabled = false;
			this.optFiltro2.ItemAppearance = appearance34;
			this.optFiltro2.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem7.DataValue = -1;
			valueListItem7.DisplayText = "Todos";
			valueListItem8.DataValue = 0;
			valueListItem8.DisplayText = "Pendiente";
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "Autorizado";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "Pagado";
			this.optFiltro2.Items.Add(valueListItem7);
			this.optFiltro2.Items.Add(valueListItem8);
			this.optFiltro2.Items.Add(valueListItem9);
			this.optFiltro2.Items.Add(valueListItem10);
			this.optFiltro2.ItemSpacingVertical = 10;
			this.optFiltro2.Location = new System.Drawing.Point(480, 72);
			this.optFiltro2.Name = "optFiltro2";
			this.optFiltro2.Size = new System.Drawing.Size(296, 24);
			this.optFiltro2.TabIndex = 759;
			this.optFiltro2.Text = "Pendiente";
			this.optFiltro2.Visible = false;
			this.optFiltro2.ValueChanged += new System.EventHandler(this.optFiltro2_ValueChanged);
			// 
			// btnCerrar
			// 
			this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCerrar.Enabled = false;
			this.btnCerrar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
			this.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrar.Location = new System.Drawing.Point(788, 127);
			this.btnCerrar.Name = "btnCerrar";
			this.btnCerrar.Size = new System.Drawing.Size(80, 23);
			this.btnCerrar.TabIndex = 760;
			this.btnCerrar.Text = "Cerrar";
			this.btnCerrar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
			// 
			// frmProvisionesFacturas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1392, 466);
			this.Controls.Add(this.btnCerrar);
			this.Controls.Add(this.optFiltro2);
			this.Controls.Add(this.optFiltro1);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.cboFecha);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnIrA);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.uGridSumas);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtIdProvision);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridProvisiones);
			this.Controls.Add(this.uGridSaldosBancarios);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmProvisionesFacturas";
			this.Text = "Provision Pagos a Proveedores";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmProvisionesFacturas_KeyDown);
			this.Load += new System.EventHandler(this.frmProvisionesFacturas_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridProvisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProvision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSaldosBancarios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSumas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iEstado = 0;
		
		private void UnloadMe()
		{
			this.Close();
		}
    private Acceso miAcceso;
		private void frmProvisionesFacturas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "044403");
			
			if (!Funcion.Permiso("916", cdsSeteoF))
			{				
				
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			DateTime Fecha = DateTime.Today;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select Fecha From Provisiones Order By Fecha", true);						
			
			while (dr.Read())
			{
				Fecha = dr.GetDateTime(0);
				this.cboFecha.Items.Add(Fecha.ToString("dd/MM/yyyy"));
			}
			dr.Close();

			this.cboFecha.Text = Fecha.ToString("dd/MM/yyyy");			

			if (this.cboFecha.Items.Count > 0)
			{
				this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoCliente, Grupo From ClienteGrupo Where Proveedor = 0 And idGrupoCliente Not In (4, 13) Order By Grupo");

				this.cboFecha.Enabled = true; 
				this.cmbGrupo.Enabled = true;
				this.txtBuscar.Enabled = true;
				this.btnVer.Enabled = true;
			}
			else
			{
				MessageBox.Show("No existen registros de Provisiones Generadas\n\nIngrese al modulo de Provision de Pagos para generar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		}

		private string EncabezadoCuenta(int iPosicion)
		{
			string sSQL = string.Format("Select SUBSTRING(b.Nombre, 1, 6) + ' ' + cc.NumCuenta From CuentaCorriente cc Inner Join Banco b On b.idBanco = cc.idBanco Where cc.Posicion = {0}", iPosicion);
			string sColumna = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			return sColumna;
		}

		private bool VisibleCuenta(int iPosicion)
		{
			string sSQL = string.Format("If Exists(Select Flujo From CuentaCorriente Where Posicion = {0}) Select Flujo From CuentaCorriente Where Posicion = {0} Else Select Convert(Bit, 0)", iPosicion);
			bool bVisible = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);
			return bVisible;
		}

		private void DiseñoGridCuentas(string sColumna, int iPosicion)
		{
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sColumna].Width = 90;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sColumna].Format = "#,##0.00";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sColumna].Header.Caption = this.EncabezadoCuenta(iPosicion);
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sColumna].CellActivation = Activation.AllowEdit;			
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sColumna].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;			
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sColumna].Hidden = !this.VisibleCuenta(iPosicion);

			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries.Add(sColumna, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sColumna], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries[sColumna].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries[sColumna].DisplayFormat = "{0: #,##0.00}";

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].Header.Caption = "...";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].CellActivation = Activation.AllowEdit;			
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].Width = 25;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].Hidden = !this.VisibleCuenta(iPosicion);
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].CellButtonAppearance.Image = Image.FromFile(@"C:\Latinium\add.png");//   \192.168.1.187\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].CellButtonAppearance.ImageHAlign = Infragistics.Win.HAlign.Center;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["btn" + sColumna].CellButtonAppearance.ImageVAlign = Infragistics.Win.VAlign.Middle;
		}

		private void DiseñoGrid ()
		{
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["idProvisionesFacturas"].Hidden = true;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["idProvisiones"].Hidden = true;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Numero"].Width = 70;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Numero"].CellActivation = Activation.ActivateOnly;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Fecha"].Width = 80;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.Disabled;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Fecha"].Header.Caption = "Emisión";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Fecha"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaRecepcion"].Width = 80;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaRecepcion"].CellActivation = Activation.Disabled;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaRecepcion"].Header.Caption = "Recepción";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaRecepcion"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Grupo"].Width = 130;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Grupo"].CellActivation = Activation.ActivateOnly;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Identificacion"].Width = 100;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Identificacion"].CellActivation = Activation.ActivateOnly;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Proveedor"].Width = 220;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Proveedor"].CellActivation = Activation.ActivateOnly;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["SaldoFactura"].Width = 90;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["SaldoFactura"].CellActivation = Activation.ActivateOnly;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["SaldoFactura"].Header.Caption = "Saldo Total";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["SaldoFactura"].Format = "#,##0.00";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["SaldoFactura"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["SaldoFactura"].Hidden = true;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["ValorPagado"].Width = 90;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["ValorPagado"].CellActivation = Activation.ActivateOnly;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["ValorPagado"].Header.Caption = "Valor Pagado";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["ValorPagado"].Format = "#,##0.00";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["ValorPagado"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["ValorPagado"].Hidden = true;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["idAutorizacionPago"].Hidden = true;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["NumPago"].Width = 60;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["NumPago"].CellActivation = Activation.Disabled;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["NumPago"].Header.Caption = "Pago N.";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["NumPago"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaPago"].Width = 90;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaPago"].CellActivation = Activation.Disabled;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaPago"].Header.Caption = "Fecha De Pago";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["FechaPago"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"].Width = 100;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"].CellActivation = Activation.Disabled;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"].Header.Caption = "Vencido";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"].Format = "#,##0.00";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries.Add("Vencido", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries["Vencido"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries["Vencido"].DisplayFormat = "{0: #,##0.00}";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"].CellAppearance.BackColor = Color.LightGray;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Vencido"].CellAppearance.BackColorDisabled = Color.LightGray;

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"].Width = 100;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"].CellActivation = Activation.Disabled;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"].Header.Caption = "Valor A Pagar";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"].Format = "#,##0.00";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries.Add("Pago", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries["Pago"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries["Pago"].DisplayFormat = "{0: #,##0.00}";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"].CellAppearance.BackColor = Color.LightBlue;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["Pago"].CellAppearance.BackColorDisabled = Color.LightBlue;
									
			for (int i = 1; i <= 20; i++)
				DiseñoGridCuentas("idCuenta" + i.ToString(), i);

			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"].Width = 100;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"].CellActivation = Activation.Disabled;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"].Header.Caption = "Total A Pagar";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"].Format = "#,##0.00";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries.Add("TotalPago", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries["TotalPago"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
			this.uGridProvisiones.DisplayLayout.Bands[0].Summaries["TotalPago"].DisplayFormat = "{0: #,##0.00}";
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"].CellAppearance.BackColor = Color.CadetBlue;
			this.uGridProvisiones.DisplayLayout.Bands[0].Columns["TotalPago"].CellAppearance.BackColorDisabled = Color.CadetBlue;

			int iContColumnas = (int)this.uGridProvisiones.DisplayLayout.Bands[0].Columns.Count;

			string sCol = "";

			for (int j = 57; j < iContColumnas; j++)
			{
				sCol = this.uGridProvisiones.DisplayLayout.Bands[0].Columns[j].Header.Caption;

				this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sCol].CellActivation = Activation.Disabled;
				this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sCol].Width = 100;					

				this.uGridProvisiones.DisplayLayout.Bands[0].Summaries.Add(sCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridProvisiones.DisplayLayout.Bands[0].Columns[sCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridProvisiones.DisplayLayout.Bands[0].Summaries[sCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridProvisiones.DisplayLayout.Bands[0].Summaries[sCol].DisplayFormat = "{0: #,##0.00}";
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Consultas()
		{
			FuncionesProcedimientos.dsGeneral(string.Format("Exec ConsultaDetalleProvisiones {0}, 4", (int)this.txtIdProvision.Value), this.uGridSaldosBancarios);

			this.uGridProvisiones.DataSource = null;

			int idGrupoCliente = 0;
			if (this.cmbGrupo.ActiveRow != null) idGrupoCliente = (int)this.cmbGrupo.Value;

			FuncionesProcedimientos.dsGeneral(string.Format("Exec ConsultaProvisionesFacturas {0}, '{1}', {2}, '{3}', {4}, {5}", 
				(int)this.txtIdProvision.Value, Convert.ToDateTime(this.cboFecha.Text).ToString("yyyyMMdd"), idGrupoCliente, this.txtBuscar.Text.ToString(), 
				(int)this.optFiltro.Value, (int)this.optFiltro1.Value), this.uGridProvisiones);
//, (int)this.optFiltro2.Value , {6}
			this.lblContador.Text = this.uGridProvisiones.Rows.Count + " REGISTROS ENCONTRADOS";

			this.DiseñoGrid();

			this.TotalesPorGrupo();					
		}

		private void TotalesPorGrupo()
		{
			FuncionesProcedimientos.dsGeneral(string.Format("Exec ProvisionesSumaGruposProveedores {0}", 
				(int)this.txtIdProvision.Value), this.uGridSumas);

			int iContColumnas = (int)this.uGridSumas.DisplayLayout.Bands[0].Columns.Count;

			string sCol = "";

			for (int j = 0; j < iContColumnas; j++)
			{
				sCol = this.uGridSumas.DisplayLayout.Bands[0].Columns[j].Header.Caption;

				this.uGridSumas.DisplayLayout.Bands[0].Columns[sCol].CellActivation = Activation.Disabled;
				this.uGridSumas.DisplayLayout.Bands[0].Columns[sCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridSumas.DisplayLayout.Bands[0].Columns[sCol].Format = "#,##0.00";

				if (j > 1) this.uGridSumas.DisplayLayout.Bands[0].Columns[sCol].Width = 90;
			
				if (j == 0)
				{
					this.uGridSumas.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.CadetBlue;
					this.uGridSumas.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColorDisabled = Color.CadetBlue;
					this.uGridSumas.DisplayLayout.Bands[0].Columns["Total"].Width = 100;
				}
			}

			sCol = "";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridSumas.Rows.All)
			{
				decimal dValor = 0.00m;

				for (int i = 0; i < iContColumnas; i++)
				{
					sCol = this.uGridSumas.DisplayLayout.Bands[0].Columns[i].Header.Caption;

					if (dr.Cells[sCol].Value != System.DBNull.Value) 
						dValor = dValor + Convert.ToDecimal(dr.Cells[sCol].Value);					
				}

				dr.Cells["Total"].Value = dValor;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{			
			this.Consultas();
		}

		private void uGridProvisiones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{	
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta1", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta2", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta3", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta4", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta5", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta6", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta7", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta8", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta9", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta10", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta11", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta12", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta13", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta14", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta15", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta16", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta17", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta18", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta19", 2);
			FuncionesProcedimientos.FormatoGrid(e, "idCuenta20", 2);
		}
		
		private void frmProvisionesFacturas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void SumaTotalPago(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			e.Cell.Row.Cells["TotalPago"].Value = 
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta1"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta2"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta3"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta4"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta5"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta6"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta7"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta8"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta9"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta10"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta11"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta12"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta13"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta14"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta15"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta16"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta17"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta18"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta19"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["idCuenta20"].Value);
		}

		private void SumatoriaPagos(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e, string sColumna, int idCuenta)
		{
			decimal dPago = 0.00m;

			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridProvisiones.Rows.All)
			//				//if ((int)dr.Cells["idCuenta"].Value == idCuenta)
			//					dPago = dPago + Convert.ToDecimal(dr.Cells[sColumna].Value);
			
			dPago = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec SumasProvisionesFacturas '{0}', {1}", sColumna, (int)this.txtIdProvision.Value));
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridSaldosBancarios.Rows.All)
			{
				if ((int)dr.Cells["idCuenta"].Value == idCuenta)
				{					
					dr.Cells["PagoProveedores"].Value = dPago;//Convert.ToDecimal(dr.Cells["PagoProveedores"].Value) + Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value);

					dr.Cells["SaldoFinal"].Value = 
						Convert.ToDecimal(dr.Cells["SaldoAFavor"].Value) - 
						Convert.ToDecimal(dr.Cells["PagoProveedores"].Value);
				}
			}
		}

		private void ActualizaValores(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e, string sColumna)
		{
			string sSQL = string.Format("Exec ProvisionesActualizaValores {0}, '{1}', {2}",
				(int)e.Cell.Row.Cells["idProvisionesFacturas"].Value, sColumna, Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value));
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		bool bActualiza = true;

		private void uGridProvisiones_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//if (!bActualiza) return;

			string sColumna = e.Cell.Column.Key;
			
			if (sColumna.Substring(0, 8) == "idCuenta")
			{
				if (e.Cell.Column.ToString() == sColumna)
				{
					bActualiza = false;
					if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value) e.Cell.Row.Cells[sColumna].Value = e.Cell.OriginalValue;
					
					if (Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value) > Convert.ToDecimal(e.Cell.Row.Cells["Pago"].Value))
					{						
						e.Cell.Row.Cells[sColumna].Value = e.Cell.OriginalValue;
						MessageBox.Show(string.Format("El valor a pagar no puede ser superior al saldo actual {0}", Convert.ToDecimal(e.Cell.Row.Cells["Pago"].Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						bActualiza = true;
						return;
					}
					
					if (Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value) > 0.00m)
					{
						if (Convert.ToDecimal(e.Cell.Row.Cells["TotalPago"].Value) > 0.00m)
						{
							e.Cell.Row.Cells[sColumna].Value = 0.00m;
							MessageBox.Show("Ya tiene asignado el pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							bActualiza = true;
							return;
						}
					}

					SumaTotalPago(sender, e);
					
					ActualizaValores(sender, e, sColumna);
					
					int iLength = 1;
					if (sColumna.Length > 9) iLength = 2;
					
					string sSQL = string.Format("Select idCuenta From CuentaCorriente Where Posicion = {0}", sColumna.Substring(8, iLength));
					int idCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					
					SumatoriaPagos(sender, e, sColumna, idCuenta);
					
					this.TotalesPorGrupo();

					bActualiza = true;
				}
			}
			
			if (e.Cell.Column.ToString() == "TotalPago")
			{
				bActualiza = true;
				ActualizaValores(sender, e, "TotalPago");
				bActualiza = false;
			}
		}

		private void uGridProvisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridProvisiones);
		}

		private void uGridSaldosBancarios_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "PagoProveedores")
			{								
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ProvisionesSaldosBancarios Set PagoProveedores = {0} Where idProvisionesSaldosBancarios = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["PagoProveedores"].Value), (int)e.Cell.Row.Cells["idProvisionesSaldosBancarios"].Value));
			}

			if (e.Cell.Column.ToString() == "SaldoFinal")
			{								
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update ProvisionesSaldosBancarios Set SaldoFinal = {0} Where idProvisionesSaldosBancarios = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["SaldoFinal"].Value), (int)e.Cell.Row.Cells["idProvisionesSaldosBancarios"].Value));
			}
		}

		private void uGridProvisiones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			string sColumna = this.uGridProvisiones.ActiveCell.Column.Key.ToString();
			if (sColumna.Length > 7)
			{
				if (sColumna.Substring(0, 8) == "idCuenta")
					if (Convert.ToDecimal(this.uGridProvisiones.ActiveRow.Cells["TotalPago"].Value) > 0.00m && 
						Convert.ToDecimal(this.uGridProvisiones.ActiveRow.Cells[sColumna].Value) == 0.00m &&
						this.uGridProvisiones.ActiveCell.Column.Key.ToString().ToUpper() == sColumna.ToUpper()) e.Handled = true;
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridProvisiones);
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{
			if (iEstado == 2)
			{
				MessageBox.Show("La provisión esta cerrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

//			if (this.cmbGrupo.ActiveRow != null)
//			{
//				MessageBox.Show("No puede autorizar pagos de un solo grupo de proveedores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.cmbGrupo.Focus();
//				return;
//			}
//
//			if (this.txtBuscar.Text.ToString().Length > 0)
//			{
//				MessageBox.Show("Para autorizar no debe haber filtros de busqueda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtBuscar.Focus();
//				return;
//			}

			this.Consultas();

			if (this.uGridProvisiones.Rows.Count == 0)
			{
				MessageBox.Show("No existen pagos por autorizar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridProvisiones.Rows.All)
				if (Convert.ToDecimal(dr.Cells["TotalPago"].Value) > 0.00m) iCont++;

			if (iCont == 0)
			{
				MessageBox.Show("No existen valores para autorizar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Esta seguro de autorizar {0} pagos seleccionados?", iCont), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridProvisiones.Rows.All)
				{
					if (Convert.ToDecimal(dr.Cells["TotalPago"].Value) > 0.00m)
					{	
						string sColumna = "";
						int idCuenta = 0;

						for (int j = 16; j < 56; j++)
						{
							sColumna = dr.Cells[j].Column.Key.ToString();

							if (Convert.ToDecimal(dr.Cells[sColumna].Value) > 0.00m)
							{
								int iLength = 1;
								if (sColumna.Length > 9) iLength = 2;
					
								string sSQL = string.Format("Select idCuenta From CuentaCorriente Where Posicion = {0}", sColumna.Substring(8, iLength));
								idCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);								
							}
							j++;
						}
					
						#region Actualiza Autorizacion de pago
						//DateTime FechaPago = (DateTime)dr.Cells["FechaPago"].Value;
						string sSQLAP = string.Format("Exec PagosAutorizados {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', ''", 
							(int)dr.Cells["idAutorizacionPago"].Value, 
							(int)dr.Cells["idCompra"].Value, 4, 
							(decimal)dr.Cells["TotalPago"].Value, 
							DateTime.Today.ToString("yyyyMMdd"), 2, //(int)dr.Cells["idFormaPago"].Value, 
							idCuenta, "PAGO AUTORIZADO DESDE MODULO DE PROVISIONES");
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAP, true);
						#endregion Actualiza Autorizacion de pago

						#region Actualiza Estado Tipo Prestamo - Anticipo
						string sSQLActorigen = string.Format("Exec ActualizaEstadoOrigenesAutPago {0}, {1}, 1, {2}", 
							4, (int)dr.Cells["idCompra"].Value, Convert.ToDecimal(dr.Cells["TotalPago"].Value));
						Funcion.EjecutaSQL(cdsSeteoF, sSQLActorigen);
						#endregion Actualiza Estado Tipo Prestamo - Anticipo					
					}
				}

				//Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ProvisionesActualizaEstado {0}", (int)this.txtIdProvision.Value));
        
				MessageBox.Show("Pagos autorizados correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnImprimir_Click(sender, e);

				this.cboFecha_SelectedIndexChanged(sender, e);
			}
		}

		private void uGridProvisiones_DoubleClick(object sender, System.EventArgs e)
		{

		}

		private void uGridProvisiones_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (iEstado == 2)
			{
				MessageBox.Show("La provisión esta cerrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Convert.ToDecimal(e.Cell.Row.Cells["Pago"].Value) == 0.00m) 
			{
				MessageBox.Show("El documento no tiene saldo vencido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Convert.ToDecimal(e.Cell.Row.Cells["TotalPago"].Value) > 0.00m) 
			{
				MessageBox.Show("Ya tiene asignado el pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int iLong = 9;
			if (e.Cell.Column.Key.ToString().Length > 12) iLong = 10;

			string sColumna = e.Cell.Column.Key.ToString().Substring(3, iLong);
			e.Cell.Row.Cells[sColumna].Value = Convert.ToDecimal(e.Cell.Row.Cells["Pago"].Value);			
		}

		private void uGridProvisiones_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{

		}

		private void btnIrA_Click(object sender, System.EventArgs e)
		{
			this.mnuIrA.Show(this.btnIrA, new Point(1, 25));
		}

		private void mnuIrAEstadoDeCuenta_Click(object sender, System.EventArgs e)
		{
			if (this.uGridProvisiones.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			frmCuentasPorPagar CXP = new frmCuentasPorPagar(this.uGridProvisiones.ActiveRow.Cells["Identificacion"].Value.ToString());
			CXP.MdiParent = MdiParent;
			CXP.Show();
		}

		private void mnuIrADocumento_Click(object sender, System.EventArgs e)
		{
			if (this.uGridProvisiones.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros en la lista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.uGridProvisiones.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.uGridProvisiones.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				frmCompra Compra = new frmCompra((int)this.uGridProvisiones.ActiveRow.Cells["idCompra"].Value, true);
				Compra.MdiParent = MdiParent;
				Compra.Show();
			}
		}

		private void optFiltro_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consultas();		

			this.optFiltro1.Enabled = false;
			this.optFiltro2.Enabled = false;

			this.optFiltro1.Value = 0;
			//this.optFiltro2.Value = -1;
			this.btnImprimir.Enabled = true;			
			if ((int)this.optFiltro.Value == 1)
			{
				this.optFiltro1.Enabled = true;
				//this.optFiltro2.Enabled = true;
			}
		}

		private void cboFecha_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.cboFecha.Text.Length == 0) return;

			this.txtIdProvision.Value = 0;
			iEstado = 0;
			this.lblEstado.Text = "";					
			this.cmbGrupo.Value = System.DBNull.Value;

			this.btnVer.Enabled = false;
			this.btnAutorizar.Enabled = false;
			this.btnIrA.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnExcel.Enabled = false;
			this.btnActualizar.Enabled = false;
			this.btnCerrar.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridProvisiones);

			string sSQL = string.Format("Exec ConsultaProvisiones '{0}'", Convert.ToDateTime(this.cboFecha.Text).ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.txtIdProvision.Value = dr.GetInt32(0);
				iEstado = dr.GetInt32(3);
				this.lblEstado.Text = dr.GetString(4);
			}
			dr.Close();

			if (iEstado == 0) FuncionesProcedimientos.EstadoGrids(true, this.uGridProvisiones);
			else if (iEstado == 1)
			{
				Consultas();

				this.btnVer.Enabled = true;
				this.btnIrA.Enabled = true;				
				this.btnExcel.Enabled = true;
				this.btnActualizar.Enabled = true;
				this.btnCerrar.Enabled = true;

				FuncionesProcedimientos.EstadoGrids(true, this.uGridProvisiones);
			}
			else if (iEstado == 2)
			{
				Consultas();
				
				this.btnVer.Enabled = true;
				this.btnIrA.Enabled = true;
				this.btnImprimir.Enabled = true;
				this.btnExcel.Enabled = true;
			}
		}

		private void uGridProvisiones_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (iEstado == 2)
			{
				MessageBox.Show("No puede eliminar lineas de Provisiones Cerradas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}

			for (int i = 0; i < e.Rows.Length; i++)
			{
				if (Convert.ToDecimal(e.Rows[i].Cells["TotalPago"].Value) > 0.00m)
				{
					MessageBox.Show("No puede eliminar lineas con valores de pago asignados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			for (int i = 0; i < e.Rows.Length; i++)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar las lineas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					string sSQL = string.Format("Delete From ProvisionesFacturas Where idProvisionesFacturas = {0}", (int)e.Rows[i].Cells["idProvisionesFacturas"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					e.DisplayPromptMsg = false;
				}
				else e.Cancel = true;
			}
		}

		private void uGridProvisiones_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.TotalesPorGrupo();		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldProvisiones = new ParameterField ();
			ParameterDiscreteValue ValidProvisiones = new ParameterDiscreteValue ();
			paramFieldProvisiones.ParameterFieldName = "@idProvisiones";
			ValidProvisiones.Value = int.Parse(this.txtIdProvision.Value.ToString());
			paramFieldProvisiones.CurrentValues.Add (ValidProvisiones);
			paramFields.Add (paramFieldProvisiones);
				
			Reporte Reporte = new Reporte("ProvisionesImpresionPagos.rpt", "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec GeneraProvisiones {0}, '{1}', ''", 
				(int)this.txtIdProvision.Value, Convert.ToDateTime(this.cboFecha.Text).ToString("yyyyMMdd"));
			this.txtIdProvision.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			MessageBox.Show("Provisión Actualizada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			Consultas();
		}

		private void cmbGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			//this.Consultas();		
		}

		private void cmbGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.Consultas();
		}

		private void cmbGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void cboFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbGrupo.Focus();
		}

		private void optFiltro1_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consultas();		

			this.btnAutorizar.Enabled = false;

			if ((int)this.optFiltro1.Value == 1)
				if (iEstado == 1) this.btnAutorizar.Enabled = true;
			//this.optFiltro2.Value = -1;
		}

		private void optFiltro2_ValueChanged(object sender, System.EventArgs e)
		{			
			this.Consultas();

			this.btnAutorizar.Enabled = false;

			if ((int)this.optFiltro2.Value == 0)
				if (iEstado == 1) this.btnAutorizar.Enabled = true;
		}

		private void btnCerrar_Click(object sender, System.EventArgs e)
		{
			if (iEstado == 2)
			{
				MessageBox.Show("La provisión esta cerrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de cerrar esta Provisión?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ProvisionesActualizaEstado {0}", (int)this.txtIdProvision.Value));
        
				MessageBox.Show("Provisión Cerrada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}
	}
}
