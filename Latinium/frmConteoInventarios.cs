using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConteoInventarios.
	/// </summary>
	public class frmConteoInventarios : System.Windows.Forms.Form
	{
		private System.Windows.Forms.DateTimePicker txtFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdLocales;
		private System.Windows.Forms.Button cmbExcel;
		private System.Windows.Forms.Button cmdBuscar;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdInvn;
		private System.ComponentModel.IContainer components;

		public frmConteoInventarios()
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

		C1.Data.C1DataRow drSeteoF;

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Selct");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConteoInventarios));
			this.grdInvn = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtFecha = new System.Windows.Forms.DateTimePicker();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.grdLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbExcel = new System.Windows.Forms.Button();
			this.cmdBuscar = new System.Windows.Forms.Button();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.grdInvn)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// grdInvn
			// 
			this.grdInvn.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdInvn.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdInvn.DisplayLayout.Appearance = appearance1;
			this.grdInvn.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdInvn.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdInvn.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdInvn.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdInvn.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdInvn.Location = new System.Drawing.Point(272, 8);
			this.grdInvn.Name = "grdInvn";
			this.grdInvn.Size = new System.Drawing.Size(728, 528);
			this.grdInvn.TabIndex = 189;
			this.grdInvn.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdInvn_InitializeLayout);
			// 
			// txtFecha
			// 
			this.txtFecha.CustomFormat = "";
			this.txtFecha.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFecha.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.txtFecha.Location = new System.Drawing.Point(80, 456);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtFecha.Size = new System.Drawing.Size(117, 26);
			this.txtFecha.TabIndex = 198;
			this.txtFecha.Value = new System.DateTime(2014, 6, 23, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 456);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 32);
			this.label1.TabIndex = 197;
			this.label1.Text = "Fecha Toma Inventario";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.grdLocales);
			this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(0, 8);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(264, 440);
			this.groupBox3.TabIndex = 196;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "  LOCALES COMERCIALES  ";
			// 
			// grdLocales
			// 
			this.grdLocales.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdLocales.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "-";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 227;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.grdLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdLocales.Location = new System.Drawing.Point(8, 16);
			this.grdLocales.Name = "grdLocales";
			this.grdLocales.Size = new System.Drawing.Size(248, 416);
			this.grdLocales.TabIndex = 170;
			this.toolTip1.SetToolTip(this.grdLocales, "Escoger Bodega para buscar información");
			// 
			// cmbExcel
			// 
			this.cmbExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmbExcel.Image")));
			this.cmbExcel.Location = new System.Drawing.Point(56, 496);
			this.cmbExcel.Name = "cmbExcel";
			this.cmbExcel.Size = new System.Drawing.Size(48, 40);
			this.cmbExcel.TabIndex = 195;
			this.toolTip1.SetToolTip(this.cmbExcel, "Exportar información a EXCEL");
			this.cmbExcel.Click += new System.EventHandler(this.cmbExcel_Click);
			// 
			// cmdBuscar
			// 
			this.cmdBuscar.Image = ((System.Drawing.Image)(resources.GetObject("cmdBuscar.Image")));
			this.cmdBuscar.Location = new System.Drawing.Point(8, 496);
			this.cmdBuscar.Name = "cmdBuscar";
			this.cmdBuscar.Size = new System.Drawing.Size(48, 40);
			this.cmdBuscar.TabIndex = 194;
			this.toolTip1.SetToolTip(this.cmdBuscar, "Buscar INFORMACION");
			this.cmdBuscar.Click += new System.EventHandler(this.cmdBuscar_Click);
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
			// frmConteoInventarios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1008, 541);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.cmbExcel);
			this.Controls.Add(this.cmdBuscar);
			this.Controls.Add(this.grdInvn);
			this.Name = "frmConteoInventarios";
			this.Text = "frmConteoInventarios";
			this.Load += new System.EventHandler(this.frmConteoInventarios_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdInvn)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void frmConteoInventarios_Load(object sender, System.EventArgs e)
		{
      
			this.txtFecha.Value = DateTime.Today;
			string stExecL = string.Format("Exec Bodega_Reporte 0");
			grdLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExecL);

		}

		private void cmbExcel_Click(object sender, System.EventArgs e)
		{
			if (this.grdInvn.Rows.Count == 0)			{				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				return;			}
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 			{				string stDir = MenuLatinium.stDirInicio + "\\"; 				ultraGridExcelExporter1.Export(this.grdInvn, saveFileDialog1.FileName);				System.Diagnostics.Process.Start(saveFileDialog1.FileName);			}

		}

		private void cmdBuscar_Click(object sender, System.EventArgs e)
		{
			Cursor.Current = Cursors.WaitCursor;
			
			string lc = "";
			int totColumna = 0;
			int fBusk = 0;
			DateTime FeBusk = (DateTime)this.txtFecha.Value;

			#region grlc
			//Buscar registros seleccionado en locales
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdLocales.Rows.All)
			{
				if ((bool)dr.Cells["Selct"].Value)
				{
					lc = lc + dr.Cells["Codigo"].Value;
					totColumna = 1;
				}
			}
			#endregion grlc
			
			if(lc.ToString().Equals(""))
				MessageBox.Show("Seleccione una BODEGA \n para realizar proceso de compraciones.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			else
			{
				string stExec = string.Format("Exec ConteInventario '{0}','{1}'", lc, FeBusk.ToString("yyyyMMdd"));
				grdInvn.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);	
			}
			//MsgAyuda.Wait();
			//TotlRegistros = this.grdDatLE.Rows.Count;
			Cursor.Current = Cursors.Default;

		}

		private void grdInvn_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
