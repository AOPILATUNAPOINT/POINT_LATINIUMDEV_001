using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConfiguracionUsuariosLotte.
	/// </summary>
	public class frmConfiguracionUsuariosLotte : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbUsuario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEmpresa;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbLocal;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTurno;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private System.Windows.Forms.GroupBox groupBox4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConfiguracionUsuariosLotte()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuariosCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Turno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuariosCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario", -1, "cmbUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa", -1, "cmbEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local", -1, "cmbLocal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno", -1, "cmbTurno");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTurno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTurno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Turno");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEmpresa = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTurno = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTurno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.SuspendLayout();
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cmbUsuario
			// 
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 127;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbUsuario.DisplayMember = "Usuario";
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(16, 104);
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(144, 96);
			this.cmbUsuario.TabIndex = 736;
			this.cmbUsuario.ValueMember = "idUsuario";
			this.cmbUsuario.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// cmbEmpresa
			// 
			this.cmbEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresa.DataSource = this.ultraDataSource3;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 187;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEmpresa.DisplayMember = "Empresa";
			this.cmbEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresa.Location = new System.Drawing.Point(168, 104);
			this.cmbEmpresa.Name = "cmbEmpresa";
			this.cmbEmpresa.Size = new System.Drawing.Size(144, 96);
			this.cmbEmpresa.TabIndex = 739;
			this.cmbEmpresa.ValueMember = "idEmpresa";
			this.cmbEmpresa.Visible = false;
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
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 79;
			ultraGridColumn6.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 166;
			ultraGridColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 235;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 178;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 142;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 115;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 9F;
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
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(600, 336);
			this.ultraGrid1.TabIndex = 740;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// cmbLocal
			// 
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocal.DataSource = this.ultraDataSource4;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 133;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbLocal.DisplayMember = "Local";
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(320, 104);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(144, 96);
			this.cmbLocal.TabIndex = 742;
			this.cmbLocal.ValueMember = "Bodega";
			this.cmbLocal.Visible = false;
			// 
			// cmbTurno
			// 
			this.cmbTurno.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTurno.DataSource = this.ultraDataSource5;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 113;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbTurno.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTurno.DisplayMember = "Turno";
			this.cmbTurno.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTurno.Location = new System.Drawing.Point(464, 208);
			this.cmbTurno.Name = "cmbTurno";
			this.cmbTurno.Size = new System.Drawing.Size(144, 96);
			this.cmbTurno.TabIndex = 744;
			this.cmbTurno.ValueMember = "idTurno";
			this.cmbTurno.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(88, 360);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 746;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(8, 360);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 745;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(-192, 347);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1088, 8);
			this.groupBox4.TabIndex = 747;
			this.groupBox4.TabStop = false;
			// 
			// frmConfiguracionUsuariosLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(616, 390);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbTurno);
			this.Controls.Add(this.cmbLocal);
			this.Controls.Add(this.cmbEmpresa);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.ultraGrid1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmConfiguracionUsuariosLotte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Usuarios Caja";
			this.Load += new System.EventHandler(this.frmConfiguracionUsuariosLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTurno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConfiguracionUsuariosLotte_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200805");

			if (!Funcion.Permiso("1757", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Configuracion de Usuarios Caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}

			ModificaGrilla(true);

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("ConsultaUsuariosConfiguracion"));

			this.cmbUsuario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec UsuariosInicio"));
			this.cmbEmpresa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idEmpresa, Empresa From Inicio.dbo.Empresas Where idEmpresa in (29, 34)"));
			this.cmbLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec UsuariosBodega"));
			this.cmbTurno.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idTurno, NombreTurno AS Turno From Inicio.dbo.Turnos"));
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			int ValUsuarioCaja = 0;
			ValUsuarioCaja = Funcion.iEscalarSQL(cdsSeteoF, string.Format("UsuariosCajaAutorizacion '{0}', {1}", MenuLatinium.stUsuario, MenuLatinium.stIdDB));

			if (ValUsuarioCaja == 1)
			{
				MessageBox.Show("No puede modificar la Configuracion de Usuarios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				return;
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ultraGrid1.Rows.All)
				{
					#region Validacion
					if (drr.Cells["Usuario"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione un Usuario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
						return;
					}
					else if (drr.Cells["Empresa"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione una Empresa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
						return; 
					}
					else if (drr.Cells["Local"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
						return; 
					}
					else if (drr.Cells["Fecha"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
						return; 
					}
					else if (drr.Cells["Turno"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione un Turno", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
						return; 
					}
						#endregion Validacion
					else
					{
						if ((int)drr.Cells["idUsuariosCaja"].Value == 0 || (int)drr.Cells["idUsuariosCaja"].Value > 0)
						{
							string sSQLDetalle = string.Format("Exec InsertaUsuariosTurnos {0}, {1}, {2}, {3}, '{4}', {5}", 
								drr.Cells["idUsuariosCaja"].Value,
								drr.Cells["Usuario"].Value,
								drr.Cells["Empresa"].Value,
								drr.Cells["Local"].Value,
								Convert.ToDateTime(drr.Cells["Fecha"].Value).ToString("yyyyMMdd"),
								drr.Cells["Turno"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
						}
					}
				}
				MessageBox.Show("Datos Generados Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idUsuariosCaja"].Value = 0;
		}
	}
}
