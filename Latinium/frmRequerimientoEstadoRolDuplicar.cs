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
	/// Descripción breve de frmRequerimientoEstadoRolDuplicar.
	/// </summary>
	public class frmRequerimientoEstadoRolDuplicar : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label lblAsignado;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamentoDuplicar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolEstadoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdRequerimientoRolEstadoObtenerTodos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRol;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstado;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDepartamento2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private System.Windows.Forms.Label lblRegistrosEstadoRol;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoEstadoRolDuplicar()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoEstadoRolDuplicar));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRolEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRolEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol", -1, "cmbRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento", -1, "cmbEstado", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento", -1, "cmbDepartamento2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionar", 1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.cmbDepartamentoDuplicar = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.lblRegistrosEstadoRol = new System.Windows.Forms.Label();
			this.udsRequerimientoRolEstadoObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugdRequerimientoRolEstadoObtenerTodos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbRol = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbDepartamento2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsRequerimientoRolCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamentoDuplicar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolEstadoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoRolEstadoObtenerTodos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDepartamento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance1;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataSource = this.udsRequerimientoDepartamentoCon;
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
			this.cmbDepartamento.Location = new System.Drawing.Point(184, 8);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(208, 21);
			this.cmbDepartamento.TabIndex = 1041;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn1,
																																								ultraDataColumn2});
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(16, 8);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1042;
			this.lblAsignado.Text = "* Departamento:";
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
			this.button1.Location = new System.Drawing.Point(104, 280);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1067;
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
			this.button3.Location = new System.Drawing.Point(16, 280);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(76, 23);
			this.button3.TabIndex = 1066;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// cmbDepartamentoDuplicar
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamentoDuplicar.Appearance = appearance2;
			this.cmbDepartamentoDuplicar.CausesValidation = false;
			this.cmbDepartamentoDuplicar.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamentoDuplicar.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamentoDuplicar.DataSource = this.udsRequerimientoDepartamentoCon;
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
			this.cmbDepartamentoDuplicar.Location = new System.Drawing.Point(184, 40);
			this.cmbDepartamentoDuplicar.MaxDropDownItems = 30;
			this.cmbDepartamentoDuplicar.Name = "cmbDepartamentoDuplicar";
			this.cmbDepartamentoDuplicar.Size = new System.Drawing.Size(208, 21);
			this.cmbDepartamentoDuplicar.TabIndex = 1068;
			this.cmbDepartamentoDuplicar.ValueMember = "idRequerimientoDepartamento";
			this.cmbDepartamentoDuplicar.ValueChanged += new System.EventHandler(this.cmbDepartamentoDuplicar_ValueChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(152, 23);
			this.label1.TabIndex = 1069;
			this.label1.Text = "* Departamento a Duplicar:";
			// 
			// lblRegistrosEstadoRol
			// 
			this.lblRegistrosEstadoRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistrosEstadoRol.ForeColor = System.Drawing.Color.Red;
			this.lblRegistrosEstadoRol.Location = new System.Drawing.Point(408, 40);
			this.lblRegistrosEstadoRol.Name = "lblRegistrosEstadoRol";
			this.lblRegistrosEstadoRol.Size = new System.Drawing.Size(240, 23);
			this.lblRegistrosEstadoRol.TabIndex = 1070;
			// 
			// udsRequerimientoRolEstadoObtenerTodos
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn7.DataType = typeof(int);
			this.udsRequerimientoRolEstadoObtenerTodos.Band.Columns.AddRange(new object[] {
																																											ultraDataColumn3,
																																											ultraDataColumn4,
																																											ultraDataColumn5,
																																											ultraDataColumn6,
																																											ultraDataColumn7});
			// 
			// ugdRequerimientoRolEstadoObtenerTodos
			// 
			this.ugdRequerimientoRolEstadoObtenerTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdRequerimientoRolEstadoObtenerTodos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdRequerimientoRolEstadoObtenerTodos.DataSource = this.udsRequerimientoRolEstadoObtenerTodos;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Appearance = appearance3;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 246;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.Caption = "Rol";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 245;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.Caption = "Estado";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Width = 238;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 76;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "Departamento";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 254;
			ultraGridColumn10.AutoEdit = false;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.Caption = "N°";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 21;
			ultraGridColumn11.DataType = typeof(bool);
			ultraGridColumn11.DefaultCellValue = false;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 148;
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
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.GroupByBox.Hidden = true;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ugdRequerimientoRolEstadoObtenerTodos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdRequerimientoRolEstadoObtenerTodos.FlatMode = true;
			this.ugdRequerimientoRolEstadoObtenerTodos.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.ugdRequerimientoRolEstadoObtenerTodos.Location = new System.Drawing.Point(16, 72);
			this.ugdRequerimientoRolEstadoObtenerTodos.Name = "ugdRequerimientoRolEstadoObtenerTodos";
			this.ugdRequerimientoRolEstadoObtenerTodos.Size = new System.Drawing.Size(728, 192);
			this.ugdRequerimientoRolEstadoObtenerTodos.TabIndex = 1071;
			// 
			// cmbRol
			// 
			this.cmbRol.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 180;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbRol.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRol.DisplayMember = "Descripcion";
			this.cmbRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRol.Location = new System.Drawing.Point(24, 168);
			this.cmbRol.Name = "cmbRol";
			this.cmbRol.Size = new System.Drawing.Size(72, 72);
			this.cmbRol.TabIndex = 1072;
			this.cmbRol.ValueMember = "idRequerimientoRol";
			this.cmbRol.Visible = false;
			// 
			// cmbEstado
			// 
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 180;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(104, 168);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(72, 72);
			this.cmbEstado.TabIndex = 1073;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.Visible = false;
			// 
			// cmbDepartamento2
			// 
			this.cmbDepartamento2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento2.DataSource = this.udsRequerimientoDepartamentoCon;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbDepartamento2.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbDepartamento2.DisplayMember = "Descripcion";
			this.cmbDepartamento2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento2.Location = new System.Drawing.Point(184, 168);
			this.cmbDepartamento2.Name = "cmbDepartamento2";
			this.cmbDepartamento2.Size = new System.Drawing.Size(64, 72);
			this.cmbDepartamento2.TabIndex = 1074;
			this.cmbDepartamento2.ValueMember = "idRequerimientoDepartamento";
			this.cmbDepartamento2.Visible = false;
			// 
			// udsRequerimientoRolCon
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.udsRequerimientoRolCon.Band.Columns.AddRange(new object[] {
																																			 ultraDataColumn8,
																																			 ultraDataColumn9});
			// 
			// udsEstado
			// 
			ultraDataColumn10.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn10,
																													ultraDataColumn11});
			// 
			// frmRequerimientoEstadoRolDuplicar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(768, 310);
			this.Controls.Add(this.cmbDepartamento2);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.cmbRol);
			this.Controls.Add(this.ugdRequerimientoRolEstadoObtenerTodos);
			this.Controls.Add(this.lblRegistrosEstadoRol);
			this.Controls.Add(this.cmbDepartamentoDuplicar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.cmbDepartamento);
			this.Controls.Add(this.lblAsignado);
			this.Name = "frmRequerimientoEstadoRolDuplicar";
			this.Text = "Estado - Rol (Duplicar)";
			this.Load += new System.EventHandler(this.frmRequerimientoEstadoRolDuplicar_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamentoDuplicar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolEstadoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugdRequerimientoRolEstadoObtenerTodos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoEstadoRolDuplicar_Load(object sender, System.EventArgs e)
		{
			//Consultar Departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			//Consultar Departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamentoDuplicar);
			RequerimientoRolEstadoObtenerTodos();
		}

		private void RequerimientoRolCon()
		{
			string squery = string.Format("Exec RequerimientoRolCon");
			this.cmbRol.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}
		private void RequerimientoEstadoObtenerActivos()
		{
			string squery = string.Format("Exec RequerimientoEstadoObtenerActivos");
			this.cmbEstado.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
		}
		private void RequerimientoDepartamentoCon2()
		{
			string squery = String.Format("Exec RequerimientoDepartamentoCon");
			this.cmbDepartamento2.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		private void RequerimientoRolEstadoObtenerTodos()
		{
			RequerimientoRolCon();
			RequerimientoEstadoObtenerActivos();
			RequerimientoDepartamentoCon2();
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
				string squery = string.Format("Exec RequerimientoRolEstadoObtenerTodos {0}", Departamento);
				this.ugdRequerimientoRolEstadoObtenerTodos.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint, squery);
				frmRequerimientoConfiguracion.MensajeCountConcatenar(this.lblRegistrosEstadoRol,this.ugdRequerimientoRolEstadoObtenerTodos);
			}
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento")) return;
			if (!Validacion.vbComboVacio(this.cmbDepartamentoDuplicar, "Seleccione el departamento a duplicar")) return;
			if (!Validacion.vbGridVacio(this.ugdRequerimientoRolEstadoObtenerTodos, "No existen datos en la tabla a duplicar", "idRequerimientoRol")) return;

			// Validar si al menos un checkbox está seleccionado en el grid
			bool isSelected = false;
			foreach (UltraGridRow row in this.ugdRequerimientoRolEstadoObtenerTodos.Rows)
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
			foreach (UltraGridRow row in this.ugdRequerimientoRolEstadoObtenerTodos.Rows)
			{
				if (row.Cells["Seleccionar"].Value != null && (bool)row.Cells["Seleccionar"].Value == true)
				{
					int idRequerimientoRol = (int)row.Cells["idRequerimientoRol"].Value;
					int idEstadoRequerimiento = (int)row.Cells["idEstadoRequerimiento"].Value;
					bool activo = (bool)row.Cells["Activo"].Value;

					try
					{
						// Llama al procedimiento almacenado para insertar los datos
						InsertarRequerimientoRolEstado(idRequerimientoRol, idEstadoRequerimiento, activo, idRequerimientoDepartamento);
						registrosExitosos++;
					}
					catch (Exception ex)
					{
						// Captura el error y aumenta el contador de registros fallidos
						registrosFallidos++;
						MessageBox.Show("Error al insertar el registro: ID Requerimiento Rol " + idRequerimientoRol + ". Detalles: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
		private void InsertarRequerimientoRolEstado(int idRequerimientoRol, int idEstadoRequerimiento, bool activo, int idRequerimientoDepartamento)
		{
			using (SqlConnection connection = new SqlConnection(frmRequerimiento.sconexionPoint))
			{
				SqlCommand command = new SqlCommand("RequerimientoRolEstadoInsertar", connection);
				command.CommandType = CommandType.StoredProcedure;

				command.Parameters.Add(new SqlParameter("@idRequerimientoRol", idRequerimientoRol));
				command.Parameters.Add(new SqlParameter("@idEstadoRequerimiento", idEstadoRequerimiento));
				command.Parameters.Add(new SqlParameter("@Activo", activo));
				command.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", idRequerimientoDepartamento));

				connection.Open();
				command.ExecuteNonQuery();
			}
		}
		
		private void cmbDepartamentoDuplicar_ValueChanged(object sender, System.EventArgs e)
		{
			RequerimientoRolEstadoObtenerTodos();
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
	}
}
