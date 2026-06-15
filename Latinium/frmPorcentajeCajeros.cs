using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using C1.Data;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPorcentajeCajeros.
	/// </summary>
	public class frmPorcentajeCajeros : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridBodegas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPorcentajeCajeros()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCajero");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionCobranza");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCajero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionCobranza");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtidBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridBodegas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.txtidBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(656, 8);
			this.groupBox1.TabIndex = 741;
			this.groupBox1.TabStop = false;
			// 
			// txtidBodega
			// 
			this.txtidBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidBodega.Enabled = false;
			this.txtidBodega.Location = new System.Drawing.Point(400, 8);
			this.txtidBodega.Name = "txtidBodega";
			this.txtidBodega.PromptChar = ' ';
			this.txtidBodega.Size = new System.Drawing.Size(40, 21);
			this.txtidBodega.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidBodega.TabIndex = 813;
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
			// uGridBodegas
			// 
			this.uGridBodegas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridBodegas.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			this.uGridBodegas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridBodegas.DisplayLayout.Appearance = appearance1;
			this.uGridBodegas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 63;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 75;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 153;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 323;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Porcentaje Cajero";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 119;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 25;
			ultraGridColumn7.Header.Caption = "CrediP";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 57;
			ultraGridColumn8.Header.Caption = "Com. Cobranza";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 95;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.uGridBodegas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridBodegas.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridBodegas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridBodegas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridBodegas.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridBodegas.DisplayLayout.Override.CellAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridBodegas.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBodegas.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBodegas.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBodegas.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridBodegas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridBodegas.Location = new System.Drawing.Point(8, 48);
			this.uGridBodegas.Name = "uGridBodegas";
			this.uGridBodegas.Size = new System.Drawing.Size(640, 512);
			this.uGridBodegas.TabIndex = 814;
			this.uGridBodegas.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridBodegas_CellChange);
			this.uGridBodegas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridBodegas_KeyDown);
			this.uGridBodegas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridBodegas_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn8.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// txtAnio
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnio.Appearance = appearance10;
			this.txtAnio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnio.Enabled = false;
			this.txtAnio.FormatString = "";
			this.txtAnio.Location = new System.Drawing.Point(40, 8);
			this.txtAnio.MaskInput = "nnnn";
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(96, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 816;
			this.txtAnio.ValueChanged += new System.EventHandler(this.txtAnio_ValueChanged);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 8);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(23, 15);
			this.ultraLabel4.TabIndex = 815;
			this.ultraLabel4.Text = "Año";
			// 
			// frmPorcentajeCajeros
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(656, 566);
			this.Controls.Add(this.txtAnio);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.uGridBodegas);
			this.Controls.Add(this.txtidBodega);
			this.Controls.Add(this.groupBox1);
			this.KeyPreview = true;
			this.Name = "frmPorcentajeCajeros";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Porcentaje Cajeros";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPorcentajeCajeros_KeyDown);
			this.Load += new System.EventHandler(this.frmPorcentajeCajeros_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmPorcentajeCajeros_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "052112");
			
			if (!Funcion.Permiso("1213", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Porcentaje Cajeros", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.txtAnio.Value = DateTime.Today.Year;

			CargarGrid();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		bool bActualiza = false;

		private void uGridBodegas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			if (e.Cell.Column.ToString() == "PorcentajeCajero")
			{	
				bActualiza = true;		
				if ((int) uGridBodegas.ActiveRow.Cells["IdBodega"].Value > 0)
				{								
					string sSQL = string.Format("update Bodega set PorcentajeCajero = {0} where  IdBodega = {1}" ,
						(decimal)e.Cell.Row.Cells["PorcentajeCajero"].Value,(int)e.Cell.Row.Cells["IdBodega"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					
				}	
				MessageBox.Show("Porcentaje Actualizado","Mensaje",MessageBoxButtons.OK,MessageBoxIcon.Information);
				bActualiza = false;
			}	

			
		}

		private void frmPorcentajeCajeros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridBodegas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridBodegas);
		}

		private void CargarGrid()
		{			
			FuncionesProcedimientos.dsGeneral("Exec RPPorcentajeCajero", this.uGridBodegas);		

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridBodegas.Rows.All)
				if(dr.Cells["PorcentajeCajero"].Value == System.DBNull.Value )dr.Cells["PorcentajeCajero"].Value = 0.00m;
		}

		private void txtAnio_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void uGridBodegas_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{


			if (e.Cell.Column.ToString() == "ComisionCobranza")
			{		
				if ((int) uGridBodegas.ActiveRow.Cells["IdBodega"].Value > 0)
				{		
								int ComisionCobranza = 0;
				
					this.uGridBodegas.UpdateData();	
						if ((bool)e.Cell.Row.Cells["ComisionCobranza"].Value == true) ComisionCobranza = 1;
					string sSQL = string.Format("update Bodega set ComisionCobranza = {0} where  IdBodega = {1}" ,
						ComisionCobranza,(int)e.Cell.Row.Cells["IdBodega"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					
				}	
				MessageBox.Show("Comision Cobranza Actualizado","Mensaje",MessageBoxButtons.OK,MessageBoxIcon.Information);
			}	

			if (e.Cell.Column.ToString() == "CrediPoint")
			{	
				if ((int) uGridBodegas.ActiveRow.Cells["IdBodega"].Value > 0)
				{		
						int CrediPoint = 0;
						this.uGridBodegas.UpdateData();
						if ((bool)e.Cell.Row.Cells["CrediPoint"].Value == true) CrediPoint = 1;

					string sSQL = string.Format("update Bodega set CrediPoint = {0} where  IdBodega = {1}" ,
						CrediPoint,(int)e.Cell.Row.Cells["IdBodega"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					
				}	
				MessageBox.Show("CrediPoint Actualizado","Mensaje",MessageBoxButtons.OK,MessageBoxIcon.Information);

			}	
		}
	}
}
