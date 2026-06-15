using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_PorcentajesComisiones.
	/// </summary>
	public class frmCom_PorcentajesComisiones : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_PorcentajesComisiones()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PorcentajesComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fijo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeVentas");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCarpa");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorComision");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorGrupo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PorcentajesComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fijo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_PorcentajesComisiones));
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "SeteoF";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 185;
			ultraGridColumn2.Header.Caption = "Cargo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 220;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 40;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 109;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 109;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "% Ventas";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 90;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "% Comision";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 109;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "% Carpa";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 109;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Valor Fijo";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 109;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.Header.Caption = "Por Grupo";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 70;
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
																										 ultraGridColumn11});
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance10;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8.5F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 56);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(786, 440);
			this.uGridDetalle.TabIndex = 761;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDetalle_KeyPress);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn11.DefaultValue = false;
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
																																 ultraDataColumn11});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataSource = this.ultraDataSource3;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(298, 167);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(200, 112);
			this.cmbCargos.TabIndex = 762;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
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
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(637, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 763;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(717, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 764;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(805, 8);
			this.groupBox1.TabIndex = 765;
			this.groupBox1.TabStop = false;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(544, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(80, 24);
			this.btnExcel.TabIndex = 766;
			this.btnExcel.Text = "Exportar";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmCom_PorcentajesComisiones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(802, 502);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.cmbCargos);
			this.KeyPreview = true;
			this.Name = "frmCom_PorcentajesComisiones";
			this.Text = "Porcentajes de Comisiones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_PorcentajesComisiones_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_PorcentajesComisiones_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Consulta()
		{
			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Com_ConsultaPorcentajesComisiones");
		}

		private void frmCom_PorcentajesComisiones_Load(object sender, System.EventArgs e)
		{
			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Cargo From Com_CargosDeVentas Where Comisiona = 1 Order By Orden");
			Consulta();
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "PorcentajeVentas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "PorcentajeComision", 2);
			FuncionesProcedimientos.FormatoGrid(e, "PorcentajeCarpa", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ValorComision", 2);
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

				#region Porcentajes repetidos
				string sSQL = string.Format("Exec Com_ValidaPorcentajes {0}, {1}, {2}, {3}", 
					(int)dr.Cells["idCom_PorcentajesComision"].Value, (int)dr.Cells["idCargo"].Value, Convert.ToDecimal(dr.Cells["Desde"].Value), Convert.ToDecimal(dr.Cells["Hasta"].Value));

				if (Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
				{
					MessageBox.Show("Esta ingresando valores repetidos en Porcentajes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDetalle.ActiveRow = dr;
					this.uGridDetalle.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes repetidos				
			}

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{
					string sSQLDetalle = string.Format("Exec Com_GuardaPorcentajesComisiones {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}",
						(int)dr.Cells["idCom_PorcentajesComision"].Value, (int)dr.Cells["idCargo"].Value, 
						(bool)dr.Cells["Fijo"].Value, (decimal)dr.Cells["Desde"].Value, (decimal)dr.Cells["Hasta"].Value, 
						(decimal)dr.Cells["PorcentajeVentas"].Value, (decimal)dr.Cells["PorcentajeComision"].Value, (decimal)dr.Cells["PorcentajeCarpa"].Value,	(decimal)dr.Cells["ValorComision"].Value,
						(bool)dr.Cells["Diferencia"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				}

				Consulta();
			}
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCom_PorcentajesComision"].Value = 0;
			e.Row.Cells["Fijo"].Value = false;
			e.Row.Cells["Desde"].Value = 0.00m;
			e.Row.Cells["Fijo"].Value = 0.00m;
			e.Row.Cells["idCom_PorcentajesComision"].Value = 0.00m;
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["PorcentajeVentas"].Value = 0.00m;
			e.Row.Cells["PorcentajeComision"].Value = 0.00m;
			e.Row.Cells["PorcentajeCarpa"].Value = 0.00m;
			e.Row.Cells["ValorComision"].Value = 0.00m;
			e.Row.Cells["Diferencia"].Value = false;
			e.Row.Cells["PorGrupo"].Value = false;
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCom_PorcentajesComision"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCom_PorcentajesComision"].Value > 0)
						{
							string sSQL = string.Format("Delete From Com_PorcentajesComision Where idCom_PorcentajesComision = {0}", (int)e.Rows[i].Cells["idCom_PorcentajesComision"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}						
					}
				}
			}

			e.DisplayPromptMsg = false;
		}

		private void frmCom_PorcentajesComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridDetalle_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
		}
	}
}
