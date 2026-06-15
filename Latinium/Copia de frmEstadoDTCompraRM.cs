using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEstadoDTCompraRM.
	/// </summary>
	public class frmEstadoDTCompraRM : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEstados;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmEstadoDTCompraRM()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDTEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			this.uGridEstados = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.uGridEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridEstados
			// 
			this.uGridEstados.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEstados.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridEstados.DisplayLayout.Appearance = appearance1;
			this.uGridEstados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 100;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridEstados.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridEstados.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridEstados.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEstados.DisplayLayout.Override.CellAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEstados.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridEstados.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstados.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstados.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEstados.Location = new System.Drawing.Point(8, 8);
			this.uGridEstados.Name = "uGridEstados";
			this.uGridEstados.Size = new System.Drawing.Size(321, 264);
			this.uGridEstados.TabIndex = 682;
			this.uGridEstados.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridEstados_AfterRowInsert);
			this.uGridEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridEstados_KeyDown);
			this.uGridEstados.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridEstados_InitializeLayout);
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
			// btnVer
			// 
			this.btnVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(8, 280);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 684;
			this.btnVer.Text = "Aceptar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(88, 280);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 24);
			this.button1.TabIndex = 685;
			this.button1.Text = "Cancelar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// frmEstadoDTCompraRM
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(337, 310);
			this.ControlBox = false;
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridEstados);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmEstadoDTCompraRM";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Estados Mercaderia de Retiro";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmEstadoDTCompraRM_KeyDown);
			this.Load += new System.EventHandler(this.frmEstadoDTCompraRM_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmEstadoDTCompraRM_Load(object sender, System.EventArgs e)
		{
			this.uGridEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTConsultaEstados");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridEstados.Rows.All)
			{	
				string sSQL = string.Format("Exec DTGuardaEstados {0}, '{1}', {2}", 
					(int)dr.Cells["idDTEstado"].Value, dr.Cells["Descripcion"].Value.ToString(), Convert.ToDecimal(dr.Cells["Descuento"].Value));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			this.uGridEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTConsultaEstados");

			this.Close();
		}

		private void frmEstadoDTCompraRM_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridEstados_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idDTEstado"].Value = 0;
			e.Row.Cells["Descuento"].Value = 0.00m;
		}

		private void uGridEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridEstados);
		}

		private void uGridEstados_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Descuento", 2);
		}
	}
}
