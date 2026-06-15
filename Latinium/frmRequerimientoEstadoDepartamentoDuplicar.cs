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
	/// Descripción breve de frmRequerimientoEstadoDepartamentoDuplicar.
	/// </summary>
	public class frmRequerimientoEstadoDepartamentoDuplicar : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamentoDuplicar;
		private System.Windows.Forms.Label lblAsignado;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgEstadoRequerimientoDepartamentoObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstadoRequerimientoDepartamentoObtenerTodos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private System.Windows.Forms.Label lblRegistrosEstadoDepartamento;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoEstadoDepartamentoDuplicar()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoEstadoDepartamentoDuplicar));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento", -1, "cmbEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar", 2);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDepartamentoDuplicar = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.lblRegistrosEstadoDepartamento = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.udgEstadoRequerimientoDepartamentoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsEstadoRequerimientoDepartamentoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamentoDuplicar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udgEstadoRequerimientoDepartamentoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstadoRequerimientoDepartamentoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.cmbDepartamentoDuplicar);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.lblRegistrosEstadoDepartamento);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 80);
			this.groupBox1.TabIndex = 1082;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// cmbDepartamento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance1;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 218;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
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
			// cmbDepartamentoDuplicar
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamentoDuplicar.Appearance = appearance2;
			this.cmbDepartamentoDuplicar.CausesValidation = false;
			this.cmbDepartamentoDuplicar.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamentoDuplicar.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Descripción";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 218;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbDepartamentoDuplicar.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
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
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(8, 16);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1072;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// lblRegistrosEstadoDepartamento
			// 
			this.lblRegistrosEstadoDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosEstadoDepartamento.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosEstadoDepartamento.Location = new System.Drawing.Point(384, 48);
			this.lblRegistrosEstadoDepartamento.Name = "lblRegistrosEstadoDepartamento";
			this.lblRegistrosEstadoDepartamento.Size = new System.Drawing.Size(240, 23);
			this.lblRegistrosEstadoDepartamento.TabIndex = 1077;
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
			this.button1.TabIndex = 1081;
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
			this.button3.TabIndex = 1080;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// udgEstadoRequerimientoDepartamentoObtenerTodos
			// 
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Appearance = appearance3;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 160;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.Caption = "Estado";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 401;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.Caption = "Departamento";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 291;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "Posición";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 156;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "Estado";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 562;
			ultraGridColumn10.AutoEdit = false;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.Caption = "N°";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 21;
			ultraGridColumn11.DataType = typeof(bool);
			ultraGridColumn11.DefaultCellValue = false;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn11.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand3.Header.Caption = "Detalle";
			ultraGridBand3.HeaderVisible = true;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.FlatMode = true;
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Location = new System.Drawing.Point(8, 96);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Name = "udgEstadoRequerimientoDepartamentoObtenerTodos";
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.Size = new System.Drawing.Size(728, 240);
			this.udgEstadoRequerimientoDepartamentoObtenerTodos.TabIndex = 1083;
			// 
			// udsEstadoRequerimientoDepartamentoObtenerTodos
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			this.udsEstadoRequerimientoDepartamentoObtenerTodos.Band.Columns.AddRange(new object[] {
																																															 ultraDataColumn1,
																																															 ultraDataColumn2,
																																															 ultraDataColumn3,
																																															 ultraDataColumn4});
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn5,
																																								ultraDataColumn6});
			// 
			// cmbEstado
			// 
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 180;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(40, 216);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(72, 72);
			this.cmbEstado.TabIndex = 1084;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.Visible = false;
			// 
			// udsEstado
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn7,
																													ultraDataColumn8});
			// 
			// frmRequerimientoEstadoDepartamentoDuplicar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(744, 350);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.udgEstadoRequerimientoDepartamentoObtenerTodos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Name = "frmRequerimientoEstadoDepartamentoDuplicar";
			this.Text = "Estado - Departamento (Duplicar)";
			this.Load += new System.EventHandler(this.frmRequerimientoEstadoDepartamentoDuplicar_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamentoDuplicar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udgEstadoRequerimientoDepartamentoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstadoRequerimientoDepartamentoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento")) return;
			if (!Validacion.vbComboVacio(this.cmbDepartamentoDuplicar, "Seleccione el departamento a duplicar")) return;
			if (!Validacion.vbGridVacio(this.udgEstadoRequerimientoDepartamentoObtenerTodos, "No existen datos en la tabla a duplicar", "idRequerimientoTipo")) return;

			// Validar si al menos un checkbox está seleccionado en el grid
			bool isSelected = false;
			foreach (UltraGridRow row in this.udgEstadoRequerimientoDepartamentoObtenerTodos.Rows)
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

			// Proceso de duplicación
			try
			{
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					conn.Open();

					foreach (UltraGridRow row in this.udgEstadoRequerimientoDepartamentoObtenerTodos.Rows)
					{
						if (row.Cells["Seleccionar"].Value != null && (bool)row.Cells["Seleccionar"].Value == true)
						{
							int idEstadoRequerimiento = (int)row.Cells["idEstadoRequerimiento"].Value;
							int idRequerimientoDepartamento = (int)this.cmbDepartamento.Value; // Asumiendo que el combo tiene el valor correcto
							int posicion = (int)row.Cells["Posicion"].Value; // O el valor que necesites para Posicion

							// Llamar al procedimiento almacenado para insertar
							using (SqlCommand cmd = new SqlCommand("EstadoRequerimientoDepartamentoInsertar", conn))
							{
								cmd.CommandType = CommandType.StoredProcedure;
								cmd.Parameters.Add("@idEstadoRequerimiento", SqlDbType.Int).Value = idEstadoRequerimiento;
								cmd.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = idRequerimientoDepartamento;
								cmd.Parameters.Add("@Posicion", SqlDbType.Int).Value = posicion;

								// Ejecutar el comando
								int newId = Convert.ToInt32(cmd.ExecuteScalar());

								// Puedes mostrar un mensaje o hacer algo con el nuevo ID
								Console.WriteLine("Nuevo registro insertado con ID: " + newId);
							}
						}
					}
				}

				MessageBox.Show("Registros duplicados con éxito.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.DialogResult = DialogResult.OK;
			}
			catch (SqlException sqlEx)
			{
				// Manejo de errores de SQL
				MessageBox.Show("Error al duplicar registros en la base de datos: " + sqlEx.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
				// Manejo de errores generales
				MessageBox.Show("Error al duplicar registros: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void EstadoRequerimientoDepartamentoObtenerTodos()
		{
			RequerimientoEstadoObtenerActivos();
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
				string squery = string.Format("Exec EstadoRequerimientoDepartamentoObtenerTodos {0}", Departamento);
				this.udgEstadoRequerimientoDepartamentoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
				frmRequerimientoConfiguracion.MensajeCountConcatenar(this.lblRegistrosEstadoDepartamento,this.udgEstadoRequerimientoDepartamentoObtenerTodos);
			}
		}

		private void cmbDepartamentoDuplicar_ValueChanged(object sender, System.EventArgs e)
		{
			EstadoRequerimientoDepartamentoObtenerTodos();
		}

		private void frmRequerimientoEstadoDepartamentoDuplicar_Load(object sender, System.EventArgs e)
		{
			//Consultar Departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			//Consultar Departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamentoDuplicar);
			EstadoRequerimientoDepartamentoObtenerTodos();
		}

		private void RequerimientoEstadoObtenerActivos()
		{
			string squery = string.Format("Exec RequerimientoEstadoObtenerActivos");
			this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}
	}
}
