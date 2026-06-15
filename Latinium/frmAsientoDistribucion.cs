using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAsientoDistribucion.
	/// </summary>
	public class frmAsientoDistribucion : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHaber;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDebe;
		bool bEdicion = false;

		public frmAsientoDistribucion()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtHaber = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDebe = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 57;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 103;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 177;
			ultraGridColumn4.Header.Caption = "Forma de Pago";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn4.Width = 92;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 239;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 86;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 114;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance5;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(0, 0);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(832, 240);
			this.uGridDetalle.TabIndex = 12;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDetalle_KeyPress);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 240);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(856, 8);
			this.groupBox1.TabIndex = 294;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Location = new System.Drawing.Point(88, 256);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 24);
			this.btnCancelar.TabIndex = 293;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Location = new System.Drawing.Point(8, 256);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(75, 24);
			this.btnAceptar.TabIndex = 292;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
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
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource6;
			ultraGridBand2.AddButtonCaption = "CompraFormaPago";
			ultraGridBand2.ColHeadersVisible = false;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.Caption = "";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 186;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(656, 200);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 86);
			this.cmbFormaPago.TabIndex = 437;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// txtHaber
			// 
			this.txtHaber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHaber.Appearance = appearance10;
			this.txtHaber.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHaber.Enabled = false;
			this.txtHaber.FormatString = "#,##0.00";
			this.txtHaber.Location = new System.Drawing.Point(512, 256);
			this.txtHaber.Name = "txtHaber";
			this.txtHaber.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHaber.PromptChar = ' ';
			this.txtHaber.Size = new System.Drawing.Size(90, 21);
			this.txtHaber.TabIndex = 439;
			this.txtHaber.Visible = false;
			// 
			// txtDebe
			// 
			this.txtDebe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDebe.Appearance = appearance11;
			this.txtDebe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDebe.Enabled = false;
			this.txtDebe.FormatString = "#,##0.00";
			this.txtDebe.Location = new System.Drawing.Point(416, 256);
			this.txtDebe.Name = "txtDebe";
			this.txtDebe.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDebe.PromptChar = ' ';
			this.txtDebe.Size = new System.Drawing.Size(90, 21);
			this.txtDebe.TabIndex = 438;
			this.txtDebe.Visible = false;
			// 
			// frmAsientoDistribucion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(832, 286);
			this.ControlBox = false;
			this.Controls.Add(this.txtHaber);
			this.Controls.Add(this.txtDebe);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.uGridDetalle);
			this.KeyPreview = true;
			this.MinimizeBox = false;
			this.Name = "frmAsientoDistribucion";
			this.Text = "Asiento Distribucion";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAsientoDistribucion_KeyDown);
			this.Load += new System.EventHandler(this.frmAsientoDistribucion_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		private void Listas ()
		{
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 4, 7, 13, 16, 19)");
		}
		private void frmAsientoDistribucion_Load(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
			Listas();
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);

			#region Cuentas
			if (e.KeyCode == Keys.F3)
			{	
					if (this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
					{																								
						using (BuscaCuentas Busqueda = new BuscaCuentas(this.uGridDetalle.ActiveRow.Cells["Codigo"].Value.ToString(), 1))
						{
							if (DialogResult.OK == Busqueda.ShowDialog())
							{
								this.uGridDetalle.ActiveRow.Cells["Codigo"].Value = Busqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
							}
						}
					}							
				
			}			
			#endregion Cuentas
		}

		private void uGridDetalle_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.uGridDetalle.ActiveRow.Cells["Debe"].Value != System.DBNull.Value)
				if ((decimal)this.uGridDetalle.ActiveRow.Cells["Debe"].Value > 0 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "HABER") e.Handled = true;

			if (this.uGridDetalle.ActiveRow.Cells["Haber"].Value != System.DBNull.Value)
				if ((decimal)this.uGridDetalle.ActiveRow.Cells["Haber"].Value > 0 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE") e.Handled = true;

		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Debe", 2);

			FuncionesProcedimientos.FormatoGrid(e, "Haber", 2);
//			(int)dr.Cells["idFormaPago"].Value == 0;
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Codigo
			if (e.Cell.Column.ToString() == "Codigo")
			{
				#region Busca Cuenta
				string sSQLContador = string.Format("Exec AsientoBuscaCuenta '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
				int iContador = Funcion.iEscalarSQL(cdsSeteoF, sSQLContador);

				if (iContador == 1)
				{
					string sSQLInfo = string.Format("Exec AsientoInfoCuenta '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLInfo, true);
					dr.Read();
					if (dr.HasRows)
					{
						e.Cell.Row.Cells["idCuenta"].Value = dr.GetInt32(0);
						e.Cell.Row.Cells["Descripcion"].Value = dr.GetString(2);
					}
					dr.Close();

				}
				else
				{
					using (BuscaCuentas Busqueda = new BuscaCuentas(e.Cell.Row.Cells["Codigo"].Value.ToString(), 1))
					{
						if (DialogResult.OK == Busqueda.ShowDialog())
						{
							e.Cell.Row.Cells["Codigo"].Value = Busqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
						}
					}
				}
				#endregion Busca Cuenta	
			}
			#endregion Codigo

			#region Forma de Pago
			if (e.Cell.Column.ToString() == "idFormaPago")
			{
				if (e.Cell.Row.Cells["idFormaPago"].Value != System.DBNull.Value)
				{	

					if ((int)e.Cell.Row.Cells["idFormaPago"].Value > 0)
					{
						#region Valida Forma de Pago con Cuenta Bancaria
						int iCuentaBancaria = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)e.Cell.Row.Cells["idCuenta"].Value), true);
						if (iCuentaBancaria == 0) 
						{
							e.Cell.Value = System.DBNull.Value;
							MessageBox.Show("La forma de pago es para cuentas bancarias", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
						#endregion Valida Forma de Pago con Cuenta Bancaria
					}
				}
				else{
				e.Cell.Row.Cells["idFormaPago"].Value=0;
				}
			}
			#endregion Forma de Pago

			#region Debe
			if (e.Cell.Column.ToString() == "Debe")
			{
				if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value)
				{	
					if (e.Cell.Row.Cells["Debe"].Value != System.DBNull.Value)
					{	
						if ((decimal)e.Cell.Row.Cells["Debe"].Value > 0.00m)
						{
							e.Cell.Value = 0.00m;
							MessageBox.Show("Seleccione una Cuenta antes de ingresar el valor del Debe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}

			}
			#endregion Debe

			#region Haber
			if (e.Cell.Column.ToString() == "Haber")
			{
				if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value)
				{							
					if (e.Cell.Row.Cells["Haber"].Value != System.DBNull.Value)
					{
						if ((decimal)e.Cell.Row.Cells["Haber"].Value > 0.00m)
						{
							e.Cell.Value = 0.00m;
							MessageBox.Show("Seleccione una Cuenta antes de ingresar el valor del Debe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
			#endregion Haber
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
			{
				if (dr.Cells["idCuenta"].Value == System.DBNull.Value && ((decimal)dr.Cells["Debe"].Value > 0 || (decimal)dr.Cells["Haber"].Value > 0))
				{
					MessageBox.Show("Ingrese una Cuenta o Elimine la Linea", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
//				if (dr.Cells["idFormaPago"].Value == System.DBNull.Value ) 
//				{
//					(int)dr.Cells["idFormaPago"].Value == 0;
//				}
				if ((decimal)dr.Cells["Debe"].Value == 0 && (decimal)dr.Cells["Haber"].Value == 0)
				{
					MessageBox.Show("Ingrese Valor al Debe o al Haber", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}


				if (dr.Cells["Detalle"].Value == System.DBNull.Value )
				{
					MessageBox.Show("Ingrese una Detalle ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				
			}
			#endregion Validacion

			DialogResult = DialogResult.OK;
			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmAsientoDistribucion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
				e.Row.Cells["idFormaPago"].Value = 0;
		}
	}
}
