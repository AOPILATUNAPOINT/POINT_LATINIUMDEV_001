using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Columns;
using Infragistics.Win;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSabanaINventarios.
	/// </summary>
	public class frmSabanaINventarios : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label47;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSabanaInventario;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdCantidad;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optExistencia;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSabanaINventarios()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmSabanaINventarios));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.btnExportar = new System.Windows.Forms.Button();
			this.label12 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label47 = new System.Windows.Forms.Label();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbSubGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdSabanaInventario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.grdCantidad = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.optExistencia = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSabanaInventario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCantidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optExistencia)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1296, 40);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 24);
			this.btnExportar.TabIndex = 1051;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(520, 34);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(53, 16);
			this.label12.TabIndex = 1056;
			this.label12.Text = "Subgrupo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(520, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(36, 16);
			this.label2.TabIndex = 1055;
			this.label2.Text = "Grupo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(520, 58);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(36, 16);
			this.label47.TabIndex = 1057;
			this.label47.Text = "Marca";
			this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbMarca
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance1;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 88;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 237;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.Enabled = false;
			this.cmbMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(584, 56);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(256, 21);
			this.cmbMarca.TabIndex = 1058;
			this.cmbMarca.ValueMember = "idMarca";
			this.cmbMarca.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbMarca_InitializeLayout);
			// 
			// cmbSubGrupo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupo.Appearance = appearance2;
			this.cmbSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 237;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbSubGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupo.DisplayMember = "Subgrupo";
			this.cmbSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupo.Enabled = false;
			this.cmbSubGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupo.Location = new System.Drawing.Point(584, 32);
			this.cmbSubGrupo.Name = "cmbSubGrupo";
			this.cmbSubGrupo.Size = new System.Drawing.Size(256, 21);
			this.cmbSubGrupo.TabIndex = 1054;
			this.cmbSubGrupo.ValueMember = "idSubgrupo";
			this.cmbSubGrupo.ValueChanged += new System.EventHandler(this.cmbSubGrupo_ValueChanged);
			// 
			// cmbGrupo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance3;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 256;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGrupo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbGrupo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			this.cmbGrupo.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			this.cmbGrupo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.Enabled = false;
			this.cmbGrupo.Location = new System.Drawing.Point(584, 8);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(256, 21);
			this.cmbGrupo.TabIndex = 1053;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			this.cmbGrupo.ValueChanged += new System.EventHandler(this.cmbGrupo_ValueChanged);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultar.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(1296, 8);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(88, 24);
			this.btnConsultar.TabIndex = 1059;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(856, 80);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(88, 23);
			this.btnVer.TabIndex = 1060;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtFecha
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance4;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(80, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1061;
			this.dtFecha.Value = new System.DateTime(2005, 6, 8, 0, 0, 0, 0);
			// 
			// txtDescripcion
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance5;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDescripcion.Location = new System.Drawing.Point(80, 32);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(360, 22);
			this.txtDescripcion.TabIndex = 1062;
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
			// txtIdSabanaInventario
			// 
			this.txtIdSabanaInventario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSabanaInventario.Enabled = false;
			this.txtIdSabanaInventario.Location = new System.Drawing.Point(1240, 0);
			this.txtIdSabanaInventario.Name = "txtIdSabanaInventario";
			this.txtIdSabanaInventario.PromptChar = ' ';
			this.txtIdSabanaInventario.Size = new System.Drawing.Size(24, 21);
			this.txtIdSabanaInventario.TabIndex = 1063;
			this.txtIdSabanaInventario.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(520, 82);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 1065;
			this.label1.Text = "Codigo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance6;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(584, 80);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(256, 21);
			this.txtCodigo.TabIndex = 1064;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(864, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1066;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grdCantidad
			// 
			this.grdCantidad.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCantidad.Cursor = System.Windows.Forms.Cursors.Default;
			appearance7.BackColor = System.Drawing.Color.White;
			this.grdCantidad.DisplayLayout.Appearance = appearance7;
			this.grdCantidad.DisplayLayout.GroupByBox.Hidden = true;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCantidad.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.grdCantidad.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCantidad.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCantidad.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdCantidad.DisplayLayout.Override.RowAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCantidad.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCantidad.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.grdCantidad.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.grdCantidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCantidad.Location = new System.Drawing.Point(8, 128);
			this.grdCantidad.Name = "grdCantidad";
			this.grdCantidad.Size = new System.Drawing.Size(1376, 392);
			this.grdCantidad.TabIndex = 1067;
			this.grdCantidad.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdCantidad_InitializeLayout_1);
			// 
			// optExistencia
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.optExistencia.Appearance = appearance15;
			this.optExistencia.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optExistencia.CheckedIndex = 0;
			this.optExistencia.Enabled = false;
			this.optExistencia.ItemAppearance = appearance16;
			this.optExistencia.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Con Existencia";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Sin Existencia";
			this.optExistencia.Items.Add(valueListItem1);
			this.optExistencia.Items.Add(valueListItem2);
			this.optExistencia.Items.Add(valueListItem3);
			this.optExistencia.ItemSpacingVertical = 10;
			this.optExistencia.Location = new System.Drawing.Point(856, 48);
			this.optExistencia.Name = "optExistencia";
			this.optExistencia.Size = new System.Drawing.Size(280, 24);
			this.optExistencia.TabIndex = 1068;
			this.optExistencia.Text = "Todos";
			this.optExistencia.Click += new System.EventHandler(this.optExistencia_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(32, 16);
			this.label3.TabIndex = 1069;
			this.label3.Text = "Corte";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 32);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(63, 16);
			this.label4.TabIndex = 1070;
			this.label4.Text = "Descripción";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1392, 8);
			this.groupBox1.TabIndex = 1071;
			this.groupBox1.TabStop = false;
			// 
			// frmSabanaINventarios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1392, 526);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.optExistencia);
			this.Controls.Add(this.grdCantidad);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label47);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtIdSabanaInventario);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.cmbSubGrupo);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.btnExportar);
			this.Name = "frmSabanaINventarios";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Sabana Inventarios";
			this.Load += new System.EventHandler(this.frmSabanaINventarios_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSabanaInventario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCantidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optExistencia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{			
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmSabanaINventarios_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0333");
			
			if (!Funcion.Permiso("2162", cdsSeteoF))
			{
				MessageBox.Show("No Tiene Acceso al Reporte Sabana de Inventarios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idGrupoArticulo, Grupo From ArticuloGrupo Where idTipoGrupo In (1, 5, 6, 13, 16) Order By Grupo");
			this.cmbSubGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idSubgrupo, SubGrupo From ArticuloSubGrupo Where Estado =1 order by SubGrupo");
			this.cmbMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " Select idMarca, Marca From ArticuloMarca  order by Marca");
		}

		private void cmbGrupo_ValueChanged(object sender, System.EventArgs e)
		{

		}

		private void cmbSubGrupo_ValueChanged(object sender, System.EventArgs e)
		{

		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.txtIdSabanaInventario.Value = 0;
			this.dtFecha.Value = null;
			this.txtDescripcion.Text = "";		
			this.grdCantidad.DataSource = null;

			this.cmbGrupo.Enabled = false;
			this.cmbSubGrupo.Enabled = false;
			this.cmbMarca.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.optExistencia.Enabled = false;
			this.btnVer.Enabled = false;
			this.btnExportar.Enabled = false;

			using (frmBuscaSabanaInv Busqueda = new frmBuscaSabanaInv())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.txtIdSabanaInventario.Value = Busqueda.GetIdSeleccionado();
					this.dtFecha.Value = Busqueda.GetFechaSeleccionada();
					this.txtDescripcion.Text = Busqueda.GetDescripcionSeleccionada();		
	
					this.cmbGrupo.Enabled = true;
					this.cmbSubGrupo.Enabled = true;
					this.cmbMarca.Enabled = true;
					this.txtCodigo.Enabled = true;
					this.btnVer.Enabled = true;
					this.btnExportar.Enabled = true;
					this.optExistencia.Enabled = true;
					
					this.btnVer_Click(sender, e);
				}
			}			
		}

		private void Buscar()
		{
			this.grdCantidad.DataSource = null;

			string sSQL = string.Format("Exec [Sabana_Consulta_Detalle] {0},{1},{2},{3},'{4}',{5}",
				(int)this.txtIdSabanaInventario.Value,
				this.cmbGrupo.Value == null ? 0 : (int)this.cmbGrupo.Value,
				this.cmbSubGrupo.Value == null ? 0 : (int)this.cmbSubGrupo.Value,
				this.cmbMarca.Value == null ? 0 : (int)this.cmbMarca.Value,
				this.txtCodigo.Text,(int)this.optExistencia.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.grdCantidad);

//			this.grdCantidad.DisplayLayout.Bands[0].SortedColumns.Add("SubGrupo", false, false);

			#region Saldo Final
			int idSubGrupo = 0;
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdCantidad.Rows.All)
				if (dr.Cells["IdSubGrupo"].Value != System.DBNull.Value)
					if (idSubGrupo == 0 || idSubGrupo != (int)dr.Cells["IdSubGrupo"].Value)
						idSubGrupo = (int)dr.Cells["IdSubGrupo"].Value;
			#endregion Saldo Final

//			this.grdCantidad.DisplayLayout.Bands[0].SortedColumns.Add("SubGrupo", false, true);
//			this.grdCantidad.Rows.ExpandAll(true);
		
			this.lblContador.Text = this.grdCantidad.Rows.Count + " REGISTROS ENCONTRADOS";

			this.Text = string.Format("Detalle de Inventario - {0} registros encontrados", this.grdCantidad.Rows.Count);
			
			DiseñoGrid();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Buscar();
		}

		private void DiseñoGrid()
		{
			if (this.grdCantidad.Rows.Count > 0)
			{				
				this.grdCantidad.DisplayLayout.Bands[0].Columns["idSabana_Inventarios_Detalle"].Hidden = true;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["idSabana_Inventarios"].Hidden = true;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["idSubGrupo"].Hidden = true;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["Grupo"].Width = 170;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Grupo"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Grupo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["SubGrupo"].Width = 170;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["SubGrupo"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["SubGrupo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["Marca"].Width = 170;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Marca"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Marca"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["Modelo"].Width = 170;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Modelo"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Modelo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["Codigo"].Width = 170;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Codigo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["Articulo"].Width = 220;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Articulo"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["Articulo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["CodigoReemplazo"].Width = 170;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["CodigoReemplazo"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["CodigoReemplazo"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_TR"].Width = 60;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_TR"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_TR"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_TR"].Header.Caption = "BI TR";
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_TR"].CellAppearance.BackColor = Color.GreenYellow;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_TR"].CellAppearance.BackColor2 = Color.GreenYellow;

				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_ST"].Width = 60;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_ST"].CellActivation = Activation.Disabled;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_ST"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_ST"].Header.Caption = "BI ST";
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_ST"].CellAppearance.BackColor = Color.Green;
				this.grdCantidad.DisplayLayout.Bands[0].Columns["BI_ST"].CellAppearance.BackColor2 = Color.Green;

				int iContador = 13;

				int iContColumnas = (int)this.grdCantidad.DisplayLayout.Bands[0].Columns.Count;

				this.grdCantidad.DisplayLayout.Bands[0].Summaries.Clear();

				string nombrCol = "";

				//int iRegion = 0 ; 

				for (int j = iContador; j < iContColumnas; j++)
				{
					nombrCol = this.grdCantidad.DisplayLayout.Bands[0].Columns[j].Header.Caption;

					if (nombrCol.Substring(3, 2) == "ST")
					{
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Green;
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor2 = Color.Green;
					}

					if (nombrCol.Substring(3, 2) == "VT")
					{
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.Yellow;
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor2 = Color.Yellow;
					}

					if (nombrCol.Substring(3, 2) == "EX")
					{
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor = Color.SkyBlue;
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.BackColor2 = Color.SkyBlue;
					}

					if (nombrCol.Substring(0, 3) == "VTA")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 150;					
					else this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 60;
					          
					if (nombrCol.Substring(nombrCol.Length - 5, 5) == "Total")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 110;

					if (nombrCol.Substring(0, 5) == "Venta")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 150;

					if (nombrCol.Substring(0, 5) == "Costo")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 120;

					if (nombrCol == "Por_Comprar")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 100;

					if (nombrCol == "Por_Comprar_USD")
					{
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 120;
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
					}
					
					if (nombrCol == "Observaciones")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 200;

					if (nombrCol == "Compra")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Width = 100;

					this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol.Replace("_", " ");

					if (nombrCol.Substring(3, 2) != "EX" || nombrCol.Substring(3, 2) != "LY" || nombrCol != "Observaciones" || nombrCol != "Compra")
						this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;

					Type colType = this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].DataType;
					if ((colType == typeof(int) || colType == typeof(decimal) || colType == typeof(double) || colType == typeof(float)))
					{
						if (!this.grdCantidad.DisplayLayout.Bands[0].Summaries.Exists(nombrCol))
						{
							this.grdCantidad.DisplayLayout.Bands[0].Summaries.Add(nombrCol,
								Infragistics.Win.UltraWinGrid.SummaryType.Sum,
								this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol],
								Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn
								);

							if (colType == typeof(int))
							{
								this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

								this.grdCantidad.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
								this.grdCantidad.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0}";
							}
							else if (colType == typeof(decimal))
							{
								this.grdCantidad.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

								this.grdCantidad.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
								this.grdCantidad.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
							} 
						}
					}
				}
			}
		}

		private void cmbMarca_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.grdCantidad.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.grdCantidad);
		}

		private void grdCantidad_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			 e.Layout.Override.RowAppearance.BackColor = System.Drawing.Color.White;
			foreach (Infragistics.Win.UltraWinGrid.UltraGridColumn column in e.Layout.Bands[0].Columns)
			{
				if (column.Key.EndsWith("_EX"))
				{
					column.CellAppearance.BackColor = System.Drawing.Color.LightBlue;
				}
				else if (column.Key.EndsWith("_VT"))
				{
					column.CellAppearance.BackColor = System.Drawing.Color.Yellow;
				}
				else if (column.Key.EndsWith("_ST"))
				{
					column.CellAppearance.BackColor = System.Drawing.Color.LightGreen;
				}
			}
		}
		private void grdCantidad_BeforeSortChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSortChangeEventArgs e)
		{
			 e.Cancel = true;
		}

		private void grdCantidad_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Override.RowAppearance.BackColor = System.Drawing.Color.White;
			foreach (Infragistics.Win.UltraWinGrid.UltraGridColumn column in e.Layout.Bands[0].Columns)
			{
				if (column.Key.EndsWith("_EX"))
				{
					column.CellAppearance.BackColor = System.Drawing.Color.LightBlue;
				}
				else if (column.Key.EndsWith("_VT"))
				{
					column.CellAppearance.BackColor = System.Drawing.Color.Yellow;
				}
				else if (column.Key.EndsWith("_ST"))
				{
					column.CellAppearance.BackColor = System.Drawing.Color.LightGreen;
				}
			}
		}

		private void optExistencia_Click(object sender, System.EventArgs e)
		{
			Buscar();
		}
	}
}
