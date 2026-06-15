using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCreacionUsuariosBodegas.
	/// </summary>
	public class frmCreacionUsuariosBodegas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.DateTimePicker dtpPeriodo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridObservaciones;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdUsuario;
		private C1.Data.C1DataSet cdsSeguridad;
		private C1.Data.C1DataSet cdsSeteo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCreacionUsuariosBodegas()
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
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCreacionUsuariosBodegas));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nidentificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCrear", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nidentificacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			this.btnVer = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.dtpPeriodo = new System.Windows.Forms.DateTimePicker();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.uGridObservaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdUsuario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(560, 45);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 739;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(280, 48);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(53, 16);
			this.label7.TabIndex = 743;
			this.label7.Text = "Vendedor";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 48);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 738;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 736;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(336, 47);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(208, 19);
			this.txtBuscar.TabIndex = 742;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(184, 14);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 740;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// dtpPeriodo
			// 
			this.dtpPeriodo.CustomFormat = "MM/yyyy";
			this.dtpPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpPeriodo.Location = new System.Drawing.Point(56, 14);
			this.dtpPeriodo.MinDate = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.ShowUpDown = true;
			this.dtpPeriodo.Size = new System.Drawing.Size(96, 20);
			this.dtpPeriodo.TabIndex = 741;
			this.dtpPeriodo.Value = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 46);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 737;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// uGridObservaciones
			// 
			this.uGridObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridObservaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridObservaciones.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridObservaciones.DisplayLayout.Appearance = appearance1;
			this.uGridObservaciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "SeteoF";
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 74;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 83;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 151;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Cargo";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 128;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 146;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance2.ImageVAlign = Infragistics.Win.VAlign.Middle;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellButtonAppearance = appearance2;
			ultraGridColumn10.Header.Caption = "...";
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 26;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridObservaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridObservaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridObservaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.CellAppearance = appearance5;
			this.uGridObservaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridObservaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridObservaciones.Location = new System.Drawing.Point(0, 96);
			this.uGridObservaciones.Name = "uGridObservaciones";
			this.uGridObservaciones.Size = new System.Drawing.Size(696, 216);
			this.uGridObservaciones.TabIndex = 744;
			this.uGridObservaciones.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridObservaciones_ClickCellButton);
			this.uGridObservaciones.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridObservaciones_AfterCellUpdate);
			this.uGridObservaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridObservaciones_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
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
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataSource = this.ultraDataSource3;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(496, 200);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(200, 112);
			this.cmbCargos.TabIndex = 745;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// txtIdUsuario
			// 
			this.txtIdUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdUsuario.Enabled = false;
			this.txtIdUsuario.Location = new System.Drawing.Point(560, 8);
			this.txtIdUsuario.MaskInput = "nn";
			this.txtIdUsuario.MinValue = 0;
			this.txtIdUsuario.Name = "txtIdUsuario";
			this.txtIdUsuario.PromptChar = ' ';
			this.txtIdUsuario.Size = new System.Drawing.Size(80, 21);
			this.txtIdUsuario.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIdUsuario.TabIndex = 748;
			this.txtIdUsuario.Visible = false;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.FillOnRequest = false;
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
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
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// frmCreacionUsuariosBodegas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(696, 310);
			this.Controls.Add(this.txtIdUsuario);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.uGridObservaciones);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.cmbBodega);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCreacionUsuariosBodegas";
			this.Text = "Usuarios Bodegas";
			this.Load += new System.EventHandler(this.frmCreacionUsuariosBodegas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCreacionUsuariosBodegas_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "020803");
			
			if (!Funcion.Permiso("2164", cdsSeteoF))
			{
				MessageBox.Show("No Tiene Acesso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.dtpPeriodo.Value = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0");			
			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 59, -1, 0");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			if (this.dtpPeriodo.Text.ToString().Length == 0) return;			

			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec Com_ConsultaVendedoresCreaUsuarios '{0}', {1}, '{2}'", 
				this.dtpPeriodo.Text.ToString(), iBodega, this.txtBuscar.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridObservaciones);
			
			this.lblContador.Text = this.uGridObservaciones.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void uGridObservaciones_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void uGridObservaciones_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnCrear")
			{
				int idCargo = Convert.ToInt32(e.Cell.Row.Cells["idCargo"].Value);
				string nUsuario = Convert.ToString(e.Cell.Row.Cells["Codigo"].Value);
				string sClave = Convert.ToString(e.Cell.Row.Cells["Nidentificacion"].Value);
				string iBodega = Convert.ToString(e.Cell.Row.Cells["Bodega"].Value);

				
				if (idCargo == 1 || idCargo == 21 )
				{
					#region Validacion
					string sSQL = string.Format("SELECT COUNT(*) FROM Usuario WHERE Nombre = '{0}'", nUsuario);
					int iContUsuario = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					if (iContUsuario > 0)
					{
						string sGrupoSQL = string.Format("SELECT idGrupo FROM Usuario WHERE Nombre = '{0}'", nUsuario);
						int iGrupo = Funcion.iEscalarSQL(cdsSeteoF, sGrupoSQL);

						if (iGrupo == 22 || iGrupo == 23)
						{
							MessageBox.Show(
								string.Format("El usuario '{0}' existe.", nUsuario),"Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Information);							
							return;
						}

						string sIdSQL = string.Format("SELECT IDUsuario FROM Usuario WHERE Nombre = '{0}'", nUsuario);
						int idUsuarioExistente = Funcion.iEscalarSQL(cdsSeteoF, sIdSQL);

						string stExec = string.Format("EXEC UsuarioCrear {0}, '{1}', '{2}', {3}, {4}, {5}",idUsuarioExistente, nUsuario, sClave, 22, 1, 1);
						this.txtIdUsuario.Value = Funcion.iEscalarSQL(cdsSeteo, stExec);

						int idUsuario = Convert.ToInt32(this.txtIdUsuario.Value);

						if (idUsuario > 0)
						{
							int nuevaBodega = Convert.ToInt32(cmbBodega.Value);
			
							string sSQLConsulta = string.Format("SELECT idUsuarioBodega FROM UsuariosBodegas WHERE idUsuario = {0}", idUsuario);
							int idUsuarioBodegaExistente = Funcion.iEscalarSQL(cdsSeteoF, sSQLConsulta);

							if (idUsuarioBodegaExistente > 0)
							{
								string sSQLEliminar = string.Format("EXEC GuardaUsuarioBodega {0}, {1}, {2}, 0",idUsuarioBodegaExistente, idUsuario, nuevaBodega);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLEliminar, true);
							}
							string sSQLInsertar = string.Format("EXEC GuardaUsuarioBodega 0, {0}, {1}, 1",idUsuario, nuevaBodega);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLInsertar, true);
						}

						#region Permisos
						string sSQLUsuarios = string.Format("Exec PermisosUsuariosCreaUsuarios '{0}', {1}", nUsuario, 1);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLUsuarios, true);
						#endregion Permisos
						MessageBox.Show(string.Format("El usuario '{0}' , fue actualizado .", nUsuario), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Information);

						return;
					}
					#endregion

					#region Crea usuario en la base de datos
					string stExec3 = string.Format("Exec UsuarioCrear {0}, '{1}', '{2}', {3}, {4}, {5}", 0, nUsuario, sClave, 22, 1, 1);
					this.txtIdUsuario.Value = Funcion.iEscalarSQL(cdsSeteo, stExec3);
					#endregion Crea usuario en la base de datos

					#region Crea usuario en el servidor
					string stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'", nUsuario);
					int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
					if (iCuentaBases == 0)
					{
						string sSQL2 = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin'; Create User {0} For Login {0}", nUsuario, "Bl45o6$9");
						Funcion.EjecutaSQL(cdsSeguridad, sSQL2);
					}
					#endregion Crea usuario en el servidor			

					#region Asignacion de bodega
					string sSQL3 = string.Format("Exec GuardaUsuarioBodega {0}, {1}, {2}, {3}", 0, (int)this.txtIdUsuario.Value, iBodega, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL3, true);
					#endregion Asignacion de bodega

					#region Permisos
					string sSQLUsuarios4 = string.Format("Exec PermisosUsuariosCreaUsuarios '{0}', {1}", nUsuario, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLUsuarios4, true);
					#endregion Permisos
					
					#region Mensaje de confirmación
					string sMensaje = "";
					sMensaje = "Usuario Registrado.";
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					#endregion Mensaje de confirmación
				}
				else if (idCargo == 3 || idCargo == 22  || idCargo == 23  || idCargo == 20 || idCargo == 2)
				{
					#region Validacion
					string sSQL = string.Format("SELECT COUNT(*) FROM Usuario WHERE Nombre = '{0}'", nUsuario);
					int iContUsuario = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					if (iContUsuario > 0)
					{
						string sGrupoSQL = string.Format("SELECT idGrupo FROM Usuario WHERE Nombre = '{0}'", nUsuario);
						int iGrupo = Funcion.iEscalarSQL(cdsSeteoF, sGrupoSQL);

						if (iGrupo == 22 || iGrupo == 23)
						{
							MessageBox.Show(
								string.Format("El usuario '{0}' existe.", nUsuario),"Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Information);							
							return;
						}

						string sIdSQL = string.Format("SELECT IDUsuario FROM Usuario WHERE Nombre = '{0}'", nUsuario);
						int idUsuarioExistente = Funcion.iEscalarSQL(cdsSeteoF, sIdSQL);

						string stExec = string.Format("EXEC UsuarioCrear {0}, '{1}', '{2}', {3}, {4}, {5}",idUsuarioExistente, nUsuario, sClave, 23, 1, 1);
						this.txtIdUsuario.Value = Funcion.iEscalarSQL(cdsSeteo, stExec);

						int idUsuario = Convert.ToInt32(this.txtIdUsuario.Value);

						if (idUsuario  > 0)
						{
							int nuevaBodega = Convert.ToInt32(cmbBodega.Value);
			
							string sSQLConsulta = string.Format("SELECT idUsuarioBodega FROM UsuariosBodegas WHERE idUsuario = {0}", idUsuario);
							int idUsuarioBodegaExistente = Funcion.iEscalarSQL(cdsSeteoF, sSQLConsulta);

							if (idUsuarioBodegaExistente > 0)
							{
								string sSQLEliminar = string.Format("EXEC GuardaUsuarioBodega {0}, {1}, {2}, 0",idUsuarioBodegaExistente, idUsuario, nuevaBodega);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLEliminar, true);
							}
							string sSQLInsertar = string.Format("EXEC GuardaUsuarioBodega 0, {0}, {1}, 1",idUsuario, nuevaBodega);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLInsertar, true);
						}

						#region Permisos
						string sSQLUsuarios = string.Format("Exec PermisosUsuariosCreaUsuarios '{0}', {1}", nUsuario, 2);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLUsuarios, true);
						#endregion Permisos
						MessageBox.Show(string.Format("El usuario '{0}' , fue actualizado .", nUsuario), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Information);

						return;
					}
					#endregion

					#region Crea usuario en la base de datos
					string stExec1 = string.Format("Exec UsuarioCrear {0}, '{1}', '{2}', {3}, {4}, {5}", 0, nUsuario, sClave, 23, 1, 1);
					this.txtIdUsuario.Value = Funcion.iEscalarSQL(cdsSeteo, stExec1);
					#endregion Crea usuario en la base de datos

					#region Crea usuario en el servidor
					string stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'", nUsuario);
					int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
					if (iCuentaBases == 0)
					{
						string sSQL2 = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin'; Create User {0} For Login {0}", nUsuario, "Bl45o6$9");
						Funcion.EjecutaSQL(cdsSeguridad, sSQL2);
					}
					#endregion Crea usuario en el servidor			

					#region Asignacion de bodega
					string sSQL3 = string.Format("Exec GuardaUsuarioBodega {0}, {1}, {2}, {3}", 0, (int)this.txtIdUsuario.Value, iBodega, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL3, true);
					#endregion Asignacion de bodega

					#region Permisos
					string sSQLUsuarios5 = string.Format("Exec PermisosUsuariosCreaUsuarios '{0}', {1}", nUsuario, 2);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLUsuarios5, true);
					#endregion Permisos
					
					#region Mensaje de confirmación
					string sMensaje = "";
					sMensaje = "Usuario Registrado.";
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					#endregion Mensaje de confirmación
				}
			}
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void uGridObservaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			 btnVer_Click(sender, e);
		}
	}
}
