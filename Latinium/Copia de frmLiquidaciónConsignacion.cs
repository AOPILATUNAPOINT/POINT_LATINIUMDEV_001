using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmLiquidaciónConsignacion.
	/// </summary>
	public class frmLiquidaciónConsignacion : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btBuscarCliente;
		private System.Windows.Forms.Label lblRUC;
		private System.Windows.Forms.Label lblNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.Misc.UltraButton btLiquidar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btSalir;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.GroupBox groupBox1;
		
		private C1DataRow drSeteoF;

		public frmLiquidaciónConsignacion()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Liquidado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Liquidar");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Liquidado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Liquidar");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.btBuscarCliente = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblRUC = new System.Windows.Forms.Label();
			this.lblNombre = new System.Windows.Forms.Label();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btLiquidar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// btBuscarCliente
			// 
			this.btBuscarCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscarCliente.Location = new System.Drawing.Point(440, 4);
			this.btBuscarCliente.Name = "btBuscarCliente";
			this.btBuscarCliente.Size = new System.Drawing.Size(88, 25);
			this.btBuscarCliente.TabIndex = 173;
			this.btBuscarCliente.Text = "Buscar";
			this.btBuscarCliente.Click += new System.EventHandler(this.btBuscarCliente_Click);
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
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 91;
			ultraGridColumn2.Header.Caption = "Código";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 223;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.Caption = "Artículo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 300;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 80;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.PromptChar = ' ';
			ultraGridColumn7.Width = 80;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(864, 320);
			this.ultraGrid1.TabIndex = 171;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(108, 16);
			this.label1.TabIndex = 170;
			this.label1.Text = "Criterio de Búsqueda";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance12;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(136, 5);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 22);
			this.txtNombre.TabIndex = 210;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
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
			// lblRUC
			// 
			this.lblRUC.AutoSize = true;
			this.lblRUC.BackColor = System.Drawing.Color.Transparent;
			this.lblRUC.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblRUC.Location = new System.Drawing.Point(8, 56);
			this.lblRUC.Name = "lblRUC";
			this.lblRUC.Size = new System.Drawing.Size(0, 20);
			this.lblRUC.TabIndex = 214;
			this.lblRUC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNombre
			// 
			this.lblNombre.AutoSize = true;
			this.lblNombre.BackColor = System.Drawing.Color.Transparent;
			this.lblNombre.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblNombre.Location = new System.Drawing.Point(288, 56);
			this.lblNombre.Name = "lblNombre";
			this.lblNombre.Size = new System.Drawing.Size(0, 20);
			this.lblNombre.TabIndex = 215;
			this.lblNombre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(816, 8);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.ReadOnly = true;
			this.txtIdCliente.Size = new System.Drawing.Size(58, 22);
			this.txtIdCliente.TabIndex = 269;
			this.txtIdCliente.Visible = false;
			// 
			// btLiquidar
			// 
			this.btLiquidar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btLiquidar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btLiquidar.Location = new System.Drawing.Point(8, 416);
			this.btLiquidar.Name = "btLiquidar";
			this.btLiquidar.Size = new System.Drawing.Size(88, 25);
			this.btLiquidar.TabIndex = 270;
			this.btLiquidar.Text = "&Liquidar";
			this.btLiquidar.Click += new System.EventHandler(this.btLiquidar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(104, 416);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 25);
			this.btCancelar.TabIndex = 271;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.Location = new System.Drawing.Point(200, 416);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(88, 25);
			this.btSalir.TabIndex = 272;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(880, 8);
			this.groupBox1.TabIndex = 273;
			this.groupBox1.TabStop = false;
			// 
			// frmLiquidaciónConsignacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(880, 446);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btLiquidar);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.lblNombre);
			this.Controls.Add(this.lblRUC);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.btBuscarCliente);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label1);
			this.Name = "frmLiquidaciónConsignacion";
			this.Text = "Liquidación de Consignaciones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmLiquidaciónConsignacion_KeyDown);
			this.Load += new System.EventHandler(this.frmLiquidaciónConsignacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmLiquidaciónConsignacion_Load(object sender, System.EventArgs e)
		{
			this.txtNombre.Select();
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscarCliente_Click(sender, e);
		}

		private void btBuscarCliente_Click(object sender, System.EventArgs e)
		{
			if (this.txtNombre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Código o el Nombre del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtNombre.Select();
				return;
			}

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select dc.idArticulo, a.Codigo, a.Articulo, 0 As Cantidad, 0 As Liquidado, 0 As Saldo, 0 As Liquidar " 
				+ "	From DetCompra dc "
				+ " Inner Join Articulo a On a.idArticulo = dc.idArticulo "
				+ " Inner Join Compra c On c.idCompra = dc.idCompra "
				+ " Where c.idCliente = 0 "
				+ " And c.idTipoFactura = 44");

			string sSQL = string.Format("Exec BuscaClientePorConsignacion '{0}'", this.txtNombre.Text.ToString());
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdCliente.Value = dr.GetInt32(0);
				this.lblRUC.Text = dr.GetString(1);
				this.lblNombre.Text = dr.GetString(2);
			}
			else
			{
				dr.Close();

				MessageBox.Show("No hay Consignaciones pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNombre.Select();
				return;
			}
			dr.Close();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				switch(e.KeyValue)
				{
					#region F3
					case (int) Keys.F3:
						#region Articulos
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
						{
							using (frmBuscarArticulosLiquidConsignacion miBusqueda = new frmBuscarArticulosLiquidConsignacion((int)this.txtIdCliente.Value))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{		
									ultraGrid1.ActiveRow.Cells["idArticulo"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
									ultraGrid1.ActiveRow.Cells["Codigo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
									ultraGrid1.ActiveRow.Cells["Articulo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Articulo"].Value.ToString();
									ultraGrid1.ActiveRow.Cells["Cantidad"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Cantidad"].Value;
									ultraGrid1.ActiveRow.Cells["Liquidado"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Liquidado"].Value;
									ultraGrid1.ActiveRow.Cells["Saldo"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Saldo"].Value;
									ultraGrid1.ActiveRow.Cells["Liquidar"].Value = 0;
								}
							}
						}
						#endregion Articulos						
						break;
						#endregion F3
					#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
					#region Flechas
					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
					#region Enter
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "");
			}
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
			{
				if (e.Cell.Column.ToString() == "Liquidar")
				{
					int dSaldo = int.Parse(e.Cell.Row.Cells["Saldo"].Value.ToString());
					int dLiquidar = int.Parse(e.Cell.Row.Cells["Liquidar"].Value.ToString());
					if (dLiquidar > dSaldo)
					{
						e.Cell.Row.Cells["Liquidar"].Value = 0;
						MessageBox.Show(string.Format("No puede liquidar mas de {0}.", dSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
				}
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.txtIdCliente.Value = 0;
			this.txtNombre.Text = "";
			this.lblRUC.Text = "";
			this.lblNombre.Text = "";

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select dc.idArticulo, a.Codigo, a.Articulo, 0 As Cantidad, 0 As Liquidado, 0 As Saldo, 0 As Liquidar " 
																																			+ "	From DetCompra dc "
																																			+ " Inner Join Articulo a On a.idArticulo = dc.idArticulo "
																																			+ " Inner Join Compra c On c.idCompra = dc.idCompra "
																																			+ " Where c.idCliente = 0 "
																																			+ " And c.idTipoFactura = 44");

			this.txtNombre.Select();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Formato Numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato Numerico

			#region Decimales para la cantidad
			stFormato = stMoneda + stCeros.Substring(0, 0);
			e.Layout.Bands[0].Columns["Liquidar"].Format = stFormato;
			
			stInput = stMascara + stNumero.Substring(0, 0);
			e.Layout.Bands[0].Columns["Liquidar"].MaskInput = stInput;
			#endregion Decimales para la cantidad
		}

		private void frmLiquidaciónConsignacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btLiquidar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if ((int)this.txtIdCliente.Value == 0)
			{
				MessageBox.Show("Seleccione un cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Select();
				return;
			}

			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los artículos a liquidar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Select();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (double.Parse(dr.Cells["Liquidar"].Value.ToString()) == 0)
				{
					MessageBox.Show(string.Format("Ingrese la Cantidad a Liquidar del Artículo '{0}'.", dr.Cells["Codigo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}
			#endregion Validacion

			if (DialogResult.No == MessageBox.Show("Antes de Liquidar los Artículos Ingresados, \n Revise si el Artículo y la Cantidad corresponden a la liquidación que desea crear. \n\n ¿Crear liquidación de consignación con estos datos.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
			#region Liquidacion
			this.Cursor = Cursors.WaitCursor;
			this.Enabled = false;
			try
			{
				#region Encabezado
				string sSQLEnc = string.Format("Exec CSLiquidaCreaEncFactura {0}", (int)this.txtIdCliente.Value);
				int idCompra = Funcion.iEscalarSQL(cdsSeteoF, sSQLEnc);
				#endregion Encabezado

				#region Detalle
				string sSQLDet = "";
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDetalle in ultraGrid1.Rows.All)
				{
					sSQLDet = string.Format("Exec CSLiquidaCreaDetalleFactura {0}, {1}, {2}, {3}", idCompra, (int)drDetalle.Cells["idArticulo"].Value, (int)this.txtIdCliente.Value, double.Parse(drDetalle.Cells["Liquidar"].Value.ToString()));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDet, true);
				}
				#endregion Detalle
				
				#region Totales
				string sSQLTotales = string.Format("Exec CSLiqTotalizaFactura {0}" ,idCompra);
				decimal dTotal = Funcion.decEscalarSQL(cdsSeteoF, sSQLTotales);
				#endregion Totales

				#region Genera Cuentas por Cobrar
				string sSQL1 = string.Format("Exec GeneraPlanDePagosCredito {0}", idCompra); 
				Funcion.EjecutaSQL(cdsSeteoF, sSQL1, true);
				#endregion Genera Cuentas por Cobrar

				#region Actualiza Estados de Consignacion
				string sSQLEstado = string.Format("Exec EstadoConsignacion {0}", (int)this.txtIdCliente.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLEstado, true);
				#endregion Actualiza Estados de Consignacion

				#region Actualiza Grid
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select dc.idArticulo, a.Codigo, a.Articulo, 0 As Cantidad, 0 As Liquidado, 0 As Saldo, 0 As Liquidar " 
																																			+ "	From DetCompra dc "
																																			+ " Inner Join Articulo a On a.idArticulo = dc.idArticulo "
																																			+ " Inner Join Compra c On c.idCompra = dc.idCompra "
																																			+ " Where c.idCliente = 0 "
																																			+ " And c.idTipoFactura = 44");
				#endregion Actualiza Grid

				MessageBox.Show("Liquidación Generada satisfactoriamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			this.Enabled = true;
			this.Cursor = Cursors.Default;
			#endregion Liquidacion
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{			
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtNombre.Text.ToString(), 44, 6, 4, 0, true))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select dc.idArticulo, a.Codigo, a.Articulo, 0 As Cantidad, 0 As Liquidado, 0 As Saldo, 0 As Liquidar " 
							+ "	From DetCompra dc "
							+ " Inner Join Articulo a On a.idArticulo = dc.idArticulo "
							+ " Inner Join Compra c On c.idCompra = dc.idCompra "
							+ " Where c.idCliente = 0 "
							+ " And c.idTipoFactura = 44");

						this.txtIdCliente.Value = miBusqueda.idCliente;
						this.lblRUC.Text = miBusqueda.Ruc;
						this.lblNombre.Text = miBusqueda.Nombre;						
					}
				}
			}
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

