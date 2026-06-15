using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoTipoDuplicar.
	/// </summary>
	public class frmRequerimientoTipoDuplicar : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamentoDuplicar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label lblAsignado;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoTipoObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoTipoObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private System.Windows.Forms.Label lblRegistrosTipoRequerimiento;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoTipoDuplicar()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoTipoDuplicar));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SeleccionDefault");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar", 2);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SeleccionDefault");
			this.cmbDepartamentoDuplicar = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.udgRequerimientoTipoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoTipoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblRegistrosTipoRequerimiento = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamentoDuplicar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoTipoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipoObtenerTodos)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// cmbDepartamentoDuplicar
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamentoDuplicar.Appearance = appearance1;
			this.cmbDepartamentoDuplicar.CausesValidation = false;
			this.cmbDepartamentoDuplicar.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamentoDuplicar.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamentoDuplicar.DataSource = this.udsRequerimientoDepartamentoCon;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 218;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbDepartamentoDuplicar.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDepartamentoDuplicar.DisplayMember = "Descripcion";
			this.cmbDepartamentoDuplicar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamentoDuplicar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamentoDuplicar.Location = new System.Drawing.Point(168, 48);
			this.cmbDepartamentoDuplicar.MaxDropDownItems = 30;
			this.cmbDepartamentoDuplicar.Name = "cmbDepartamentoDuplicar";
			this.cmbDepartamentoDuplicar.Size = new System.Drawing.Size(208, 21);
			this.cmbDepartamentoDuplicar.TabIndex = 1075;
			this.cmbDepartamentoDuplicar.ValueMember = "idRequerimientoDepartamento";
			this.cmbDepartamentoDuplicar.ValueChanged += new System.EventHandler(this.cmbDepartamentoDuplicar_ValueChanged);
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn1,
																																								ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(152, 23);
			this.label1.TabIndex = 1076;
			this.label1.Text = "* Departamento a Duplicar:";
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(656, 48);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1074;
			this.button1.Text = "Salir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(656, 16);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(80, 23);
			this.button3.TabIndex = 1073;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// cmbDepartamento
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance2;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataSource = this.udsRequerimientoDepartamentoCon;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Descripción";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 218;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(168, 16);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(208, 21);
			this.cmbDepartamento.TabIndex = 1071;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(8, 16);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1072;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// udgRequerimientoTipoObtenerTodos
			// 
			this.udgRequerimientoTipoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoTipoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoTipoObtenerTodos.DataSource = this.udsRequerimientoTipoObtenerTodos;
			appearance3.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Appearance = appearance3;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 87;
			ultraGridColumn6.AutoEdit = false;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 230;
			ultraGridColumn7.AutoEdit = false;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 110;
			ultraGridColumn8.AutoEdit = false;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "Departamento";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 206;
			ultraGridColumn9.AutoEdit = false;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "Seleccion default";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 234;
			ultraGridColumn10.AutoEdit = false;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.Caption = "Estado";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 488;
			ultraGridColumn11.AutoEdit = false;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Header.Caption = "N°";
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 21;
			ultraGridColumn12.DataType = typeof(bool);
			ultraGridColumn12.DefaultCellValue = true;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Nullable = Infragistics.Win.UltraWinGrid.Nullable.Nothing;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn12.Width = 133;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			ultraGridBand3.Header.Caption = "Detalle";
			ultraGridBand3.HeaderVisible = true;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.udgRequerimientoTipoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgRequerimientoTipoObtenerTodos.FlatMode = true;
			this.udgRequerimientoTipoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgRequerimientoTipoObtenerTodos.Location = new System.Drawing.Point(8, 96);
			this.udgRequerimientoTipoObtenerTodos.Name = "udgRequerimientoTipoObtenerTodos";
			this.udgRequerimientoTipoObtenerTodos.Size = new System.Drawing.Size(728, 200);
			this.udgRequerimientoTipoObtenerTodos.TabIndex = 1070;
			// 
			// udsRequerimientoTipoObtenerTodos
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			this.udsRequerimientoTipoObtenerTodos.Band.Columns.AddRange(new object[] {
																																								 ultraDataColumn3,
																																								 ultraDataColumn4,
																																								 ultraDataColumn5,
																																								 ultraDataColumn6,
																																								 ultraDataColumn7});
			// 
			// lblRegistrosTipoRequerimiento
			// 
			this.lblRegistrosTipoRequerimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosTipoRequerimiento.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosTipoRequerimiento.Location = new System.Drawing.Point(384, 48);
			this.lblRegistrosTipoRequerimiento.Name = "lblRegistrosTipoRequerimiento";
			this.lblRegistrosTipoRequerimiento.Size = new System.Drawing.Size(240, 23);
			this.lblRegistrosTipoRequerimiento.TabIndex = 1077;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.cmbDepartamentoDuplicar);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.lblRegistrosTipoRequerimiento);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 80);
			this.groupBox1.TabIndex = 1078;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// frmRequerimientoTipoDuplicar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(744, 310);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.udgRequerimientoTipoObtenerTodos);
			this.Name = "frmRequerimientoTipoDuplicar";
			this.Text = "Tipo Requerimiento (Duplicar)";
			this.Load += new System.EventHandler(this.frmRequerimientoTipoDuplicar_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamentoDuplicar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoTipoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoTipoObtenerTodos)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoTipoDuplicar_Load(object sender, System.EventArgs e)
		{
			//Consultar Departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			//Consultar Departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamentoDuplicar);
			RequerimientoTipoObtenerActivos();
		}

		private void RequerimientoTipoObtenerActivos()
		{
			int Departamento = 0;
			try
			{
				Departamento = (int)this.cmbDepartamentoDuplicar.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			if(Departamento>0)
			{
				string squery = string.Format("Exec RequerimientoTipoObtenerActivos {0}", Departamento);
				this.udgRequerimientoTipoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
				frmRequerimientoConfiguracion.MensajeCountConcatenar(this.lblRegistrosTipoRequerimiento, this.udgRequerimientoTipoObtenerTodos);
			}
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento")) return;
			if (!Validacion.vbComboVacio(this.cmbDepartamentoDuplicar, "Seleccione el departamento a duplicar")) return;
			if (!Validacion.vbGridVacio(this.udgRequerimientoTipoObtenerTodos, "No existen datos en la tabla a duplicar", "idRequerimientoTipo")) return;

			// Validar si al menos un checkbox está seleccionado en el grid
			bool isSelected = false;
			foreach (UltraGridRow row in this.udgRequerimientoTipoObtenerTodos.Rows)
			{
				if (row.Cells["Seleccionar"].Value != null && (bool)row.Cells["Seleccionar"].Value == true)
				{
					isSelected = true;
					break;
				}
			}

			if (!isSelected)
			{
				MessageBox.Show("Debe seleccionar al menos un registro para duplicar.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			#endregion Validación

			#region Guardar en base de datos
			int idRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			int registrosExitosos = 0;
			int registrosFallidos = 0;

			// Recorre el grid nuevamente para insertar los registros seleccionados
			foreach (UltraGridRow row in this.udgRequerimientoTipoObtenerTodos.Rows) // Cambié 'var' por 'UltraGridRow'
			{
				if (row.Cells["Seleccionar"].Value != null && (bool)row.Cells["Seleccionar"].Value == true)
				{
					string descripcion = row.Cells["Descripcion"].Value.ToString();
					bool estado = (bool)row.Cells["Estado"].Value;
					bool seleccionDefault = (bool)row.Cells["SeleccionDefault"].Value;

					try
					{
						// Llama al procedimiento almacenado para insertar los datos
						InsertarRequerimientoTipo(descripcion, estado, idRequerimientoDepartamento, seleccionDefault);
						registrosExitosos++;
					}
					catch (Exception ex)
					{
						// Captura el error y aumenta el contador de registros fallidos
						registrosFallidos++;
						MessageBox.Show("Error al insertar el registro: " + descripcion + ". Detalles: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}

			// Mostrar mensajes de éxito o error según los resultados
			if (registrosExitosos > 0 && registrosFallidos == 0)
			{
				MessageBox.Show("Se insertaron correctamente " + registrosExitosos + " registros.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.DialogResult = DialogResult.OK;
			}
			else if (registrosExitosos > 0 && registrosFallidos > 0)
			{
				MessageBox.Show("Se insertaron correctamente " + registrosExitosos + " registros, pero fallaron " + registrosFallidos + ".", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			else if (registrosFallidos > 0)
			{
				MessageBox.Show("Todos los registros fallaron. Ningún dato fue insertado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			#endregion Guardar en base de datos
		}

		// Método auxiliar para llamar al procedimiento almacenado
		private void InsertarRequerimientoTipo(string descripcion, bool estado, int idRequerimientoDepartamento, bool seleccionDefault)
		{
			using (SqlConnection connection = new SqlConnection(frmRequerimiento.sconexionPoint))
			{
				SqlCommand command = new SqlCommand("RequerimientoTipoInsertar", connection);
				command.CommandType = CommandType.StoredProcedure;

				command.Parameters.Add(new SqlParameter("@Descripcion", descripcion));
				command.Parameters.Add(new SqlParameter("@Estado", estado));
				command.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", idRequerimientoDepartamento));
				command.Parameters.Add(new SqlParameter("@SeleccionDefault", seleccionDefault));

				connection.Open();
				command.ExecuteNonQuery();
			}
		}

		private void cmbDepartamentoDuplicar_ValueChanged(object sender, System.EventArgs e)
		{
			RequerimientoTipoObtenerActivos();
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}


	}
}
