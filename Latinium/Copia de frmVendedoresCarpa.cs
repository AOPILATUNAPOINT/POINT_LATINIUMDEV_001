using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmVendedoresCarpa.
	/// </summary>
	public class frmVendedoresCarpa : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		public System.Windows.Forms.ComboBox cmbAño;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmVendedoresCarpa()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbAño = new System.Windows.Forms.ComboBox();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(216, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 725;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbAño
			// 
			this.cmbAño.Location = new System.Drawing.Point(64, 40);
			this.cmbAño.MaxDropDownItems = 6;
			this.cmbAño.Name = "cmbAño";
			this.cmbAño.Size = new System.Drawing.Size(96, 21);
			this.cmbAño.TabIndex = 723;
			this.cmbAño.SelectedValueChanged += new System.EventHandler(this.cmbAño_SelectedValueChanged);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 40);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(42, 15);
			this.ultraLabel4.TabIndex = 724;
			this.ultraLabel4.Text = "Periodo";
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance4;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 80);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1000, 168);
			this.uGridDetalle.TabIndex = 722;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
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
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 728;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 50;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(224, 21);
			this.cmbBodega.TabIndex = 727;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 57;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedores.DisplayMember = "Codigo";
			this.cmbVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbVendedores.Location = new System.Drawing.Point(264, 96);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(200, 112);
			this.cmbVendedores.TabIndex = 729;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// frmVendedoresCarpa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1016, 254);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.cmbAño);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.cmbVendedores);
			this.Name = "frmVendedoresCarpa";
			this.Text = "frmVendedoresCarpa";
			this.Load += new System.EventHandler(this.frmVendedoresCarpa_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmVendedoresCarpa_Load(object sender, System.EventArgs e)
		{
			this.cmbAño.Items.Add("08/2016");
			this.cmbAño.Items.Add("09/2016");
			this.cmbAño.Items.Add("10/2016");
			this.cmbAño.Items.Add("11/2016");
			this.cmbAño.Items.Add("12/2016");
			this.cmbAño.Items.Add("01/2017");
			this.cmbAño.Items.Add("02/2017");
			this.cmbAño.Items.Add("03/2017");
			
			//FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 And Credito = 1 Order By Nombre");
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.uGridDetalle.DataSource = null;

			string sSQL = string.Format("Exec ConsultaVendedorCarpa '{0}', {1}", this.cmbAño.Text.ToString(), (int)this.cmbBodega.Value);
			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.DiseñoGrid();
		}

		private void DiseñoGrid()
		{
			try
			{
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["Periodo"].Hidden = true;
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
							
				int iContColumnas = (int)this.uGridDetalle.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 2;

				string nombrCol = "";

				for (int j = iContador; j < iContColumnas; j++)
				{
					nombrCol = this.uGridDetalle.DisplayLayout.Bands[0].Columns[j].Header.Caption;
				
					this.uGridDetalle.DisplayLayout.Bands[0].Columns[nombrCol].Width = 150;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol.Substring(6, 2);
					this.uGridDetalle.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.DropDown;
					this.uGridDetalle.DisplayLayout.Bands[0].Columns[nombrCol].ValueList = this.cmbVendedores;
				}     			
			
				this.uGridDetalle.DisplayLayout.Bands[0].SummaryFooterCaption = "TOTALES";	
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void cmbAño_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;
			DateTime dtFecha = DateTime.Parse("01/" + this.cmbAño.Text.ToString()); 
			int iUltimoDia = FuncionesProcedimientos.iUltimoDiaDelMes(dtFecha.Month, dtFecha.Year);
			dtFecha = DateTime.Parse(iUltimoDia.ToString() + "/" + this.cmbAño.Text.ToString());
			this.cmbVendedores.DataSource = FuncionesProcedimientos.Lista_Vendedores(dtFecha, (int)this.cmbBodega.Value, cdsSeteoF); // Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ListaVendedores '{0}'", this.cmbAño.Text.ToString()));
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);
		}

		private void GuardaVendedorCarpa(string sPeriodo, int iBodega, DateTime Fecha, int idPersonal)
		{
			string sSQL = string.Format("Exec GuardaVendedorCarpa '{0}', {1}, '{2}', {3}", sPeriodo, iBodega, Fecha.ToString("yyyyMMdd"), idPersonal);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}
		
		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{					
		//@Periodo Varchar(7), @Bodega Int, @Fecha Datetime, @idPersonal Int
			string sNombreColumna = e.Cell.Column.Key;
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.ToString() == sNombreColumna)
			{
				DateTime dtFecha = DateTime.Parse(sColumna + "/" + this.cmbAño.Text.ToString());				

				GuardaVendedorCarpa(this.cmbAño.Text.ToString(), (int)e.Cell.Row.Cells["Bodega"].Value, dtFecha, (int)e.Cell.Row.Cells[sNombreColumna].Value);
			}
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["Bodega"].Value = (int) this.cmbBodega.Value;
		}		
	}
}
