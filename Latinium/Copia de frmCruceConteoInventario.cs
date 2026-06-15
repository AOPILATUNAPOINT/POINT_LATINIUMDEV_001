using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCruceConteoInventario.
	/// </summary>
	public class frmCruceConteoInventario : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCierre;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantidad;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCruceConteoInventario()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int idArticulo = 0;
		int Bodega = 0;
		DateTime dtFecha = DateTime.Now;
		string Codigo = "";
    int Faltante = 0;
		int Cantidad = 0;
		int idLog = 0;

		public frmCruceConteoInventario(int IdLog, int IdArticulo, int IBodega, DateTime DTFecha, string sCodigo, int iFaltante, int iCantidad)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			idLog = IdLog;
			idArticulo = IdArticulo;
			Bodega = IBodega;
			dtFecha = DTFecha;
			Codigo = sCodigo;
			Faltante = iFaltante;
			Cantidad = iCantidad;

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idLogAjusteCantidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobrante");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cruce");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idLogAjusteCantidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobrante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cruce");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCruceConteoInventario));
			this.uGridCierre = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtCantidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.uGridCierre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridCierre
			// 
			this.uGridCierre.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCierre.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCierre.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridCierre.DisplayLayout.Appearance = appearance1;
			this.uGridCierre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 229;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 111;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 250;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 75;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.uGridCierre.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridCierre.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridCierre.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCierre.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCierre.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCierre.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridCierre.Location = new System.Drawing.Point(5, 88);
			this.uGridCierre.Name = "uGridCierre";
			this.uGridCierre.Size = new System.Drawing.Size(421, 192);
			this.uGridCierre.TabIndex = 189;
			this.uGridCierre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCierre_KeyDown);
			this.uGridCierre.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCierre_AfterCellUpdate);
			this.uGridCierre.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCierre_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 8);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 16);
			this.label7.TabIndex = 371;
			this.label7.Text = "Codigo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance8;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(72, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(200, 22);
			this.txtCodigo.TabIndex = 370;
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
			// txtCantidad
			// 
			this.txtCantidad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtCantidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantidad.Enabled = false;
			this.txtCantidad.Location = new System.Drawing.Point(72, 40);
			this.txtCantidad.Name = "txtCantidad";
			this.txtCantidad.PromptChar = ' ';
			this.txtCantidad.Size = new System.Drawing.Size(64, 22);
			this.txtCantidad.TabIndex = 376;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(49, 16);
			this.label1.TabIndex = 375;
			this.label1.Text = "A Cruzar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(344, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 377;
			this.btnGuardar.Text = "Cruzar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(440, 8);
			this.groupBox1.TabIndex = 378;
			this.groupBox1.TabStop = false;
			// 
			// frmCruceConteoInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(434, 288);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtCantidad);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.uGridCierre);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCruceConteoInventario";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cruce Conteo Inventario";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCruceConteoInventario_KeyDown);
			this.Load += new System.EventHandler(this.frmCruceConteoInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridCierre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCruceConteoInventario_Load(object sender, System.EventArgs e)
		{
			this.txtCodigo.Text = Codigo;			
			this.txtCantidad.Value = Cantidad;

			string sSQL = string.Format("Exec CierreConteoCruce {0}, {1}, '{2}'", idArticulo, Bodega, dtFecha.ToString("yyyyMMdd"));
			this.uGridCierre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL); 	
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmCruceConteoInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void uGridCierre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (uGridCierre.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{							
							#region Flechas
						case (int) Keys.Up:
							//							if (uGridCierre.ActiveCell.DroppedDown) return;

							uGridCierre.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridCierre.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							uGridCierre.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							uGridCierre.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridCierre.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							uGridCierre.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							//							if (uGridCierre.ActiveCell.DroppedDown) return;
			
							uGridCierre.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridCierre.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							uGridCierre.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							uGridCierre.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridCierre.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							uGridCierre.PerformAction(UltraGridAction.EnterEditMode, false, false);
							if(uGridCierre.ActiveCell.DroppedDown == false) uGridCierre.ActiveCell.DroppedDown = true;
							//						uGridCierre.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}
	
		private void uGridCierre_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Cruce")
			{
				if ((int)e.Cell.Row.Cells["Cruce"].Value > (int)e.Cell.Row.Cells["Sobrante"].Value)
				{
					e.Cell.Value = 0;
					MessageBox.Show("La Cantidad a Cruzar no puede ser Mayor al Valor del Sobrante", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}

				int iSuma = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCierre.Rows.All)
				{
					iSuma = iSuma + (int)dr.Cells["Cruce"].Value;
				}
	
				if (iSuma > (int)this.txtCantidad.Value)
				{
					e.Cell.Value = 0;
					MessageBox.Show(string.Format("Maximo a Cruzar es {0}", (int)this.txtCantidad.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iSuma = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCierre.Rows.All)
			{
				iSuma = iSuma + (int)dr.Cells["Cruce"].Value;
			}
	
			if (iSuma == 0)
			{				
				MessageBox.Show("Escriba la Cantidad a Cruzar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCierre.Rows.All)
				{
					if ((int)dr.Cells["Cruce"].Value > 0)
					{
						string sSQL = string.Format("Exec GuardaCrucesInventario {0}, {1}, {2}, {3}, '{4}'", 
							idLog, (int)this.txtCantidad.Value, (int)dr.Cells["idLogAjusteCantidades"].Value, (int)dr.Cells["Cruce"].Value, "");
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaLogAjusteCantidades {0}, '{1}', {2}, {3}", 
							(int)dr.Cells["Cruce"].Value, dr.Cells["Codigo"].Value.ToString(), (int)dr.Cells["idLogAjusteCantidades"].Value, idLog));

//						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update LogAjusteCantidades Set NotasDelSistema = ISNULL(NotasDelSistema, '') + '{0}' Where idLogAjusteCantidades = {1}", 
//							dr.Cells["Codigo"].Value.ToString(), idLog));
					}
				}

				this.DialogResult = DialogResult.OK;
			}
		}

		private void uGridCierre_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

