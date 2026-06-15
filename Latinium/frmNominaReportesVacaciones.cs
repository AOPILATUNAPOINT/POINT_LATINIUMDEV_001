using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmNominaReportesVacaciones.
	/// </summary>
	public class frmNominaReportesVacaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNominaReportesVacaciones()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		private Acceso miAcceso;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNominaReportesVacaciones));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellidos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroDeCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salida");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Obligatorio");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tomados");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pendientes");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroDeCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Obligatorio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tomados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pendientes");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(800, 12);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 703;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1480, 8);
			this.groupBox1.TabIndex = 702;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(528, 9);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 699;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(436, 9);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 698;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// txtBuscar
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance1;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(120, 10);
			this.txtBuscar.MaxLength = 300;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(296, 22);
			this.txtBuscar.TabIndex = 697;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 12);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(82, 17);
			this.label2.TabIndex = 701;
			this.label2.Text = "Cedula/Nombre";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 83;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 97;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 154;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 153;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Centro De Costo";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 91;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 98;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 106;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 143;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 91;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 91;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 91;
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
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDetalle.DisplayLayout.GroupByBox.Hidden = true;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
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
			appearance11.FontData.SizeInPoints = 10F;
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
			this.uGridDetalle.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 56);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1246, 304);
			this.uGridDetalle.TabIndex = 700;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
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
			// optEstados
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance15;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance16;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Activos";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Inactivos";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(620, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(160, 24);
			this.optEstados.TabIndex = 706;
			this.optEstados.Text = "Activos";
			// 
			// frmNominaReportesVacaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1262, 366);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.uGridDetalle);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmNominaReportesVacaciones";
			this.Text = "Nomina - Reportes Vacaciones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNominaReportesVacaciones_KeyDown);
			this.Load += new System.EventHandler(this.frmNominaReportesVacaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}


		private void frmNominaReportesVacaciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200311");

			if (!Funcion.Permiso("1769", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExportar.Enabled = true;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmNominaReportesVacaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if ((int)this.optEstados.Value==1)
			{
				string sSQL = string.Format("Exec ReporteCalculaDiasVacaciones '{0}'", this.txtBuscar.Text.ToString());
				this.uGridDetalle.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				this.lblContador.Text = this.uGridDetalle.Rows.Count + " REGISTROS ENCONTRADOS";
			}
			else if ((int)this.optEstados.Value==2)
			{
				string sSQL = string.Format("Exec ReporteCalculaDiasVacacionesInactivos '{0}'", this.txtBuscar.Text.ToString());
				this.uGridDetalle.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				this.lblContador.Text = this.uGridDetalle.Rows.Count + " REGISTROS ENCONTRADOS";
			}
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}
	}
}
