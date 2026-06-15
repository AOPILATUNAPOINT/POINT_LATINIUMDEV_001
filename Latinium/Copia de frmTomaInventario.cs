using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTomaInventario.
	/// </summary>
	public class frmTomaInventario : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button cmbExcel;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblBodega;
		private System.ComponentModel.IContainer components;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdDTomas;
		private System.Windows.Forms.GroupBox optGroup;
		private System.Windows.Forms.RadioButton rdbCont;
		private System.Windows.Forms.RadioButton rdbInven;
		private Infragistics.Win.UltraWinGrid.UltraGrid grContoInv;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCanCon;

		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
			
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo txtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		private System.Windows.Forms.Button btnRegistrar;
		private System.Windows.Forms.Timer timer1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscaSerial;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCierre;

		int iBodega = 0;
		string sBodega = "";
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optBodegaExhibicion;
		private System.Windows.Forms.Button btnVerToma;
		DateTime dtFecha = DateTime.Today;
		private System.Windows.Forms.Label lblMsjeContExist;
		int iExistencia = 0;
		
		public frmTomaInventario(int IBodega, DateTime Fecha)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			iBodega = IBodega;
			dtFecha = Fecha;
			
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTomaInventario));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NomArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambioArticulo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NomArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CambioArticulo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.cmbExcel = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.grdDTomas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnVerToma = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.optGroup = new System.Windows.Forms.GroupBox();
			this.rdbInven = new System.Windows.Forms.RadioButton();
			this.rdbCont = new System.Windows.Forms.RadioButton();
			this.grContoInv = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.txtCanCon = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnRegistrar = new System.Windows.Forms.Button();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.txtBuscaSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtIdCierre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optBodegaExhibicion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblMsjeContExist = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.grdDTomas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.optGroup.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grContoInv)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCanCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCierre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optBodegaExhibicion)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbExcel
			// 
			this.cmbExcel.Image = ((System.Drawing.Image)(resources.GetObject("cmbExcel.Image")));
			this.cmbExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.cmbExcel.Location = new System.Drawing.Point(168, 16);
			this.cmbExcel.Name = "cmbExcel";
			this.cmbExcel.Size = new System.Drawing.Size(80, 24);
			this.cmbExcel.TabIndex = 178;
			this.cmbExcel.Text = "Exportar";
			this.cmbExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.cmbExcel, "Exportar información EXCELL");
			this.cmbExcel.Click += new System.EventHandler(this.cmbExcel_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 17);
			this.label1.TabIndex = 180;
			this.label1.Text = "Artículo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 75);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 17);
			this.label2.TabIndex = 182;
			this.label2.Text = "Serial";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(528, 7);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 184;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// grdDTomas
			// 
			this.grdDTomas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.grdDTomas.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdDTomas.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdDTomas.DisplayLayout.Appearance = appearance1;
			this.grdDTomas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Codigo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 220;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Local";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 60;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Serial";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 219;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 80;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.grdDTomas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdDTomas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdDTomas.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Tahoma";
			appearance3.FontData.SizeInPoints = 8.5F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdDTomas.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.grdDTomas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDTomas.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDTomas.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdDTomas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdDTomas.Location = new System.Drawing.Point(8, 136);
			this.grdDTomas.Name = "grdDTomas";
			this.grdDTomas.Size = new System.Drawing.Size(520, 320);
			this.grdDTomas.TabIndex = 188;
			this.grdDTomas.AfterRowsDeleted += new System.EventHandler(this.grdDTomas_AfterRowsDeleted);
			this.grdDTomas.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.grdDTomas_BeforeRowsDeleted);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
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
			// btnVerToma
			// 
			this.btnVerToma.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVerToma.Image = ((System.Drawing.Image)(resources.GetObject("btnVerToma.Image")));
			this.btnVerToma.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVerToma.Location = new System.Drawing.Point(416, 72);
			this.btnVerToma.Name = "btnVerToma";
			this.btnVerToma.Size = new System.Drawing.Size(80, 23);
			this.btnVerToma.TabIndex = 195;
			this.btnVerToma.Text = "Ver Toma";
			this.btnVerToma.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnVerToma, "Buscar información...!!!");
			this.btnVerToma.Click += new System.EventHandler(this.btnVerToma_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 17);
			this.label4.TabIndex = 191;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.ForeColor = System.Drawing.Color.Firebrick;
			this.lblBodega.Location = new System.Drawing.Point(232, 7);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(0, 22);
			this.lblBodega.TabIndex = 194;
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// optGroup
			// 
			this.optGroup.Controls.Add(this.rdbInven);
			this.optGroup.Controls.Add(this.rdbCont);
			this.optGroup.Controls.Add(this.cmbExcel);
			this.optGroup.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.optGroup.Location = new System.Drawing.Point(672, 8);
			this.optGroup.Name = "optGroup";
			this.optGroup.Size = new System.Drawing.Size(256, 48);
			this.optGroup.TabIndex = 197;
			this.optGroup.TabStop = false;
			this.optGroup.Text = "Exportar";
			// 
			// rdbInven
			// 
			this.rdbInven.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rdbInven.Location = new System.Drawing.Point(80, 16);
			this.rdbInven.Name = "rdbInven";
			this.rdbInven.Size = new System.Drawing.Size(80, 24);
			this.rdbInven.TabIndex = 1;
			this.rdbInven.Text = "Inventario";
			this.rdbInven.Click += new System.EventHandler(this.rdbInven_Click);
			// 
			// rdbCont
			// 
			this.rdbCont.Checked = true;
			this.rdbCont.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rdbCont.Location = new System.Drawing.Point(8, 16);
			this.rdbCont.Name = "rdbCont";
			this.rdbCont.Size = new System.Drawing.Size(64, 24);
			this.rdbCont.TabIndex = 0;
			this.rdbCont.TabStop = true;
			this.rdbCont.Text = "Conteo";
			this.rdbCont.Click += new System.EventHandler(this.rdbCont_Click);
			// 
			// grContoInv
			// 
			this.grContoInv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grContoInv.Cursor = System.Windows.Forms.Cursors.Default;
			this.grContoInv.DataSource = this.ultraDataSource1;
			appearance6.BackColor = System.Drawing.Color.White;
			this.grContoInv.DisplayLayout.Appearance = appearance6;
			this.grContoInv.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 248;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Local";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 53;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance7;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance8;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Width = 80;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.grContoInv.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.grContoInv.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grContoInv.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.grContoInv.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Tahoma";
			appearance10.FontData.SizeInPoints = 8.5F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grContoInv.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.grContoInv.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grContoInv.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grContoInv.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.grContoInv.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grContoInv.Location = new System.Drawing.Point(536, 136);
			this.grContoInv.Name = "grContoInv";
			this.grContoInv.Size = new System.Drawing.Size(472, 320);
			this.grContoInv.TabIndex = 198;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(328, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(48, 17);
			this.label5.TabIndex = 200;
			this.label5.Text = "Cantidad";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCanCon
			// 
			appearance13.BorderColor = System.Drawing.SystemColors.ActiveBorder;
			appearance13.BorderColor3DBase = System.Drawing.Color.Gray;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCanCon.Appearance = appearance13;
			this.txtCanCon.AutoSize = false;
			this.txtCanCon.BorderStyle = Infragistics.Win.UIElementBorderStyle.Dotted;
			this.txtCanCon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCanCon.Enabled = false;
			this.txtCanCon.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCanCon.Location = new System.Drawing.Point(416, 40);
			this.txtCanCon.MinValue = 0;
			this.txtCanCon.Name = "txtCanCon";
			this.txtCanCon.PromptChar = ' ';
			this.txtCanCon.Size = new System.Drawing.Size(80, 22);
			this.txtCanCon.TabIndex = 3;
			this.txtCanCon.Click += new System.EventHandler(this.txtCanCon_Click);
			this.txtCanCon.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCanCon_KeyPress);
			this.txtCanCon.Enter += new System.EventHandler(this.txtCanCon_Enter);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 120);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1016, 8);
			this.groupBox1.TabIndex = 237;
			this.groupBox1.TabStop = false;
			// 
			// txtFecha
			// 
			appearance14.FontData.Name = "Tahoma";
			appearance14.FontData.SizeInPoints = 8.25F;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFecha.Appearance = appearance14;
			this.txtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.txtFecha.DateButtons.Add(dateButton1);
			this.txtFecha.Enabled = false;
			this.txtFecha.Format = "dd/MM/yyyy HH:mm";
			this.txtFecha.Location = new System.Drawing.Point(64, 8);
			this.txtFecha.Name = "txtFecha";
			this.txtFecha.NonAutoSizeHeight = 23;
			this.txtFecha.Size = new System.Drawing.Size(152, 21);
			this.txtFecha.SpinButtonsVisible = true;
			this.txtFecha.TabIndex = 0;
			this.txtFecha.Value = ((object)(resources.GetObject("txtFecha.Value")));
			// 
			// txtCodigo
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance15;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCodigo.Location = new System.Drawing.Point(64, 40);
			this.txtCodigo.MaxLength = 37;
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(250, 22);
			this.txtCodigo.TabIndex = 2;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCodigo_KeyPress);
			this.txtCodigo.Validated += new System.EventHandler(this.txtCodigo_Validated);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtSerial
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance16;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSerial.Location = new System.Drawing.Point(64, 72);
			this.txtSerial.MaxLength = 37;
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(250, 22);
			this.txtSerial.TabIndex = 4;
			this.txtSerial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerial_KeyPress);
			this.txtSerial.ValueChanged += new System.EventHandler(this.txtSerial_ValueChanged);
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.Location = new System.Drawing.Point(632, 32);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(32, 21);
			this.txtIdArticulo.TabIndex = 324;
			this.txtIdArticulo.Visible = false;
			// 
			// btnRegistrar
			// 
			this.btnRegistrar.Enabled = false;
			this.btnRegistrar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnRegistrar.Image = ((System.Drawing.Image)(resources.GetObject("btnRegistrar.Image")));
			this.btnRegistrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRegistrar.Location = new System.Drawing.Point(328, 72);
			this.btnRegistrar.Name = "btnRegistrar";
			this.btnRegistrar.Size = new System.Drawing.Size(80, 23);
			this.btnRegistrar.TabIndex = 325;
			this.btnRegistrar.Text = "&Registrar";
			this.btnRegistrar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRegistrar.Click += new System.EventHandler(this.btnRegistrar_Click);
			// 
			// timer1
			// 
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// txtBuscaSerial
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscaSerial.Appearance = appearance17;
			this.txtBuscaSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscaSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscaSerial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBuscaSerial.Location = new System.Drawing.Point(736, 72);
			this.txtBuscaSerial.MaxLength = 37;
			this.txtBuscaSerial.Name = "txtBuscaSerial";
			this.txtBuscaSerial.Size = new System.Drawing.Size(192, 22);
			this.txtBuscaSerial.TabIndex = 326;
			this.txtBuscaSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscaSerial_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(600, 75);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(120, 17);
			this.label6.TabIndex = 327;
			this.label6.Text = "Busca Serial / Producto";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdCierre
			// 
			this.txtIdCierre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCierre.Enabled = false;
			this.txtIdCierre.Location = new System.Drawing.Point(632, 8);
			this.txtIdCierre.Name = "txtIdCierre";
			this.txtIdCierre.PromptChar = ' ';
			this.txtIdCierre.Size = new System.Drawing.Size(32, 21);
			this.txtIdCierre.TabIndex = 328;
			this.txtIdCierre.Visible = false;
			// 
			// optBodegaExhibicion
			// 
			appearance18.FontData.Name = "Tahoma";
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.optBodegaExhibicion.Appearance = appearance18;
			this.optBodegaExhibicion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optBodegaExhibicion.CheckedIndex = 0;
			this.optBodegaExhibicion.ItemAppearance = appearance19;
			this.optBodegaExhibicion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = false;
			valueListItem1.DisplayText = "Bodega";
			valueListItem2.DataValue = true;
			valueListItem2.DisplayText = "Exhibición";
			this.optBodegaExhibicion.Items.Add(valueListItem1);
			this.optBodegaExhibicion.Items.Add(valueListItem2);
			this.optBodegaExhibicion.ItemSpacingVertical = 10;
			this.optBodegaExhibicion.Location = new System.Drawing.Point(344, 6);
			this.optBodegaExhibicion.Name = "optBodegaExhibicion";
			this.optBodegaExhibicion.Size = new System.Drawing.Size(152, 24);
			this.optBodegaExhibicion.TabIndex = 612;
			this.optBodegaExhibicion.Text = "Bodega";
			// 
			// lblMsjeContExist
			// 
			this.lblMsjeContExist.AutoSize = true;
			this.lblMsjeContExist.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMsjeContExist.ForeColor = System.Drawing.Color.Firebrick;
			this.lblMsjeContExist.Location = new System.Drawing.Point(64, 104);
			this.lblMsjeContExist.Name = "lblMsjeContExist";
			this.lblMsjeContExist.Size = new System.Drawing.Size(0, 16);
			this.lblMsjeContExist.TabIndex = 613;
			this.lblMsjeContExist.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmTomaInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1016, 464);
			this.Controls.Add(this.lblMsjeContExist);
			this.Controls.Add(this.optBodegaExhibicion);
			this.Controls.Add(this.txtIdCierre);
			this.Controls.Add(this.txtBuscaSerial);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.btnRegistrar);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtFecha);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtCanCon);
			this.Controls.Add(this.grContoInv);
			this.Controls.Add(this.optGroup);
			this.Controls.Add(this.btnVerToma);
			this.Controls.Add(this.grdDTomas);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmTomaInventario";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Toma de Inventario";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTomaInventario_KeyDown);
			this.Load += new System.EventHandler(this.frmTomaInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdDTomas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.optGroup.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grContoInv)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCanCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCierre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optBodegaExhibicion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void frmTomaInventario_Load(object sender, System.EventArgs e)
		{
			this.txtFecha.Value = dtFecha;     
			string sSQLGuardaCierre = string.Format("Exec GuardaCierreConteoInventario {0}, '{1}'", iBodega, dtFecha.ToString("yyyyMMdd"));
			this.txtIdCierre.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGuardaCierre, true);

			string sSQL = string.Format("Select Codigo, Nombre From Bodega Where Bodega = {0}", iBodega);
			SqlDataReader drBodega = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drBodega.Read();
			if (drBodega.HasRows)
			{
				sBodega = drBodega.GetString(0);
				this.Text = "Toma de Inventario: " + drBodega.GetString(1);				
			}
			drBodega.Close();

			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From CierreConteoInventario Where Bodega = '{0}' And CONVERT(DATE, Fecha) = '{1}'", iBodega, dtFecha.ToString("yyyyMMdd")));
										
			if (iEstado == 1)
			{
				this.txtCanCon.Enabled = true;
				this.txtCodigo.Enabled = true;
				this.txtSerial.Enabled = true;

				this.btnRegistrar.Enabled = true;
				
				this.lblEstado.Text = "ABIERTO";
			}
			else this.lblEstado.Text = "CERRADO";			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbExcel_Click(object sender, System.EventArgs e)
		{			
			if(this.rdbCont.Checked) FuncionesProcedimientos.ExportaExcel(this.grdDTomas);

			if(this.rdbInven.Checked) FuncionesProcedimientos.ExportaExcel(this.grContoInv);
		}

		private void rdbCont_Click(object sender, System.EventArgs e)
		{
			if (this.rdbCont.Checked) this.rdbInven.Checked = false;
		}

		private void rdbInven_Click(object sender, System.EventArgs e)
		{
			if (this.rdbInven.Checked) this.rdbCont.Checked = false;
		}
		
		private void grdDTomas_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No != MessageBox.Show("¿Está Seguro de Eliminar esta Serial del Conteo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				for (int i=0; i<e.Rows.Length; i++)
				{					
					string sSQL = string.Format("Exec Elimina_Registro_Conteo {0}", (int)e.Rows[i].Cells["idC"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
					MessageBox.Show("Serial Eliminada del Conteo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				e.DisplayPromptMsg = false;
			}
			else e.Cancel = true;			
		}

		private void txtCanCon_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtSerial.Focus();
		}

		private void grdDTomas_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.btnVerToma_Click(sender,e);
		}

		private void txtCanCon_Click(object sender, System.EventArgs e)
		{
			this.txtCanCon.SelectAll();
		}

		private void txtCanCon_Enter(object sender, System.EventArgs e)
		{
			this.txtCanCon.SelectAll();
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaArticulos Busqueda = new BuscaArticulos(this.txtCodigo.Text, 1, 1, 0, Convert.ToDateTime(this.txtFecha.Value), false))
				{
					if (DialogResult.OK == Busqueda.ShowDialog()) 
					{
						this.txtCodigo.Text = Busqueda.sCodigo;
						this.txtSerial.Focus();
					}
				}
			}
		}

		private void txtCodigo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtSerial.Focus();
		}

		private void txtSerial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsControl(e.KeyChar)) e.Handled = false;
			else if (Char.IsLetterOrDigit(e.KeyChar)) e.Handled = false; 				
			else e.Handled = true;

			if (e.KeyChar == 13) this.btnRegistrar_Click(sender, e);
		}

		private void ValidaManejaSerial ()
		{
			string sSQL41 = string.Format("Select dbo.Valida_ManejaSerie('{0}')", this.txtCodigo.Text.ToString());
			if (Funcion.sEscalarSQL(cdsSeteoF, sSQL41).ToString().Trim().Equals("SI"))
			{
				this.txtCanCon.Value = 0;
				this.txtCanCon.Enabled = false;
				this.txtSerial.Focus();
			}
			else
			{				
				this.txtCanCon.Enabled = true;
				this.txtCanCon.Focus();			
			}
		}

		private void txtCodigo_Validated(object sender, System.EventArgs e)
		{
			iExistencia = 0;
			this.txtIdArticulo.Value = 0;

			if (this.txtCodigo.Text.ToString().Length > 0)
			{
				int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Articulo Where idTipoGrupo In (1, 6) And (Codigo = '{0}' Or Articulo = '{0}')", this.txtCodigo.Text.ToString()));
				if (iCont == 1)
				{
					string sSQL = string.Format("Select idArticulo, Articulo From Articulo Where Codigo = '{0}'", this.txtCodigo.Text.ToString()); 
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						this.txtIdArticulo.Value = drArticulo.GetInt32(0);				
						drArticulo.Close();

						ValidaManejaSerial ();
					}
					else 
					{					
						drArticulo.Close();

						this.txtIdArticulo.Value = 0;
						this.txtSerial.Focus();
					}
				}
				else 
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtCodigo.Text, 1, 1, 0, Convert.ToDateTime(this.txtFecha.Value), false))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog()) 
						{
							this.txtCodigo.Text = miBusqueda.sCodigo;
							this.txtIdArticulo.Value = miBusqueda.idArticulo;

							ValidaManejaSerial ();

							this.txtSerial.Focus();
						}
						else this.txtIdArticulo.Value = 0;
					}
				}

				iExistencia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.CalculaExistenciasInventario('{0}', {1}, {2})", 
					Convert.ToDateTime(this.txtFecha.Value).ToString("yyyyMMdd"), (int)this.txtIdArticulo.Value, iBodega));

				//MsjeContExis();
				this.btnVerToma_Click(sender, e);
			}			
		}
		
		private void MsjeContExis()
		{
			if (this.txtCanCon.Enabled)
				this.lblMsjeContExist.Text = this.txtCanCon.Value.ToString() + " ITEMS REGISTRADOS DE " + this.iExistencia.ToString();
			else
				this.lblMsjeContExist.Text = this.grdDTomas.Rows.Count + " ITEMS REGISTRADOS DE " + this.iExistencia.ToString();
		}

		private void btnRegistrar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.txtFecha.Value;
	
			string sMensaje = "";

			int iRes = 0;
			int idSerial = 0;
			int idArticulo = 0;

			if ((int)this.txtIdArticulo.Value == 0 && this.txtSerial.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese un Articulo o una Serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}

			#region Valida En Transacciones
			int idTipoFactura = 0;
			string sTipoFactura = "";
			string sNumero = "";
			string sSQLVT = string.Format("Exec ConteoValidaSerialEnTransacciones '{0}', {1}", this.txtSerial.Text.ToString(), iBodega);
			SqlDataReader drVT = Funcion.rEscalarSQL(cdsSeteoF, sSQLVT, true);
			drVT.Read();
			if (drVT.HasRows)
			{
				idTipoFactura = drVT.GetInt32(0);
				sTipoFactura = drVT.GetString(1);
				sNumero = drVT.GetString(2);
				drVT.Close();							
			}
			drVT.Close();

			if (idTipoFactura == 2)
			{
				MessageBox.Show(string.Format("Esta Serial esta registrada en el Documento '{0}' Numero '{1}', No puede registrarla", sTipoFactura, sNumero), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
				this.txtSerial.Text = "";
				this.txtIdArticulo.Value = 0;
				this.txtCodigo.Text = "";
				this.txtSerial.Focus();
				return;
			}
			else if (idTipoFactura > 0 && DialogResult.No == MessageBox.Show(string.Format("Esta Serial esta registrada en el Documento '{0}' Numero '{1}',\n\n\n¿Seguro de Ingresarla?", sTipoFactura, sNumero), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))return;	
				
			
			#endregion Valida En Transacciones

			#region Valida Con Articulo Seleccionado
			bool bManejaSerial = false;
			if ((int)this.txtIdArticulo.Value > 0) 
			{
				bManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)this.txtIdArticulo.Value), true);
				if (!bManejaSerial)
				{
					#region No Maneja Serial
					if ((int)this.txtCanCon.Value == 0)
					{
						MessageBox.Show("Ingrese la Cantidad de este Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtCanCon.Focus();
						this.txtIdArticulo.Value = 0;
						this.txtCanCon.Value = 0;
						this.txtCodigo.Text = "";
						return;
					}
					else
					{
						string sSQLSinSerial = string.Format("Exec Registrar_Conteo_Inventario {0},'{1}','{2}','{3}','{4}','{5}',{6},{7},'{8}',{9}", 
							1, this.txtCodigo.Text.ToString(), this.txtSerial.Text.ToString(), sBodega, MenuLatinium.stUsuario, dtFecha.ToString("yyyyMMdd HH:mm"),	0, (int)this.txtCanCon.Value, dtFecha.ToString("hh:mm"),(bool) this.optBodegaExhibicion.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLSinSerial);

						this.btnVerToma_Click(sender, e);

						this.txtSerial.Text = "";
						this.txtIdArticulo.Value = 0;
						this.txtCodigo.Text = "";
						this.txtCanCon.Value = 0;

						return;
					}
					#endregion No Maneja Serial
				}
			
				if (bManejaSerial)
				{
					#region Maneja Serial
					sMensaje = "";
					string sSQL = string.Format("Exec ValidaSerialEnConteo {0}, '{1}', '{2}', {3}", (int)this.txtIdArticulo.Value, this.txtSerial.Text.ToString(), dtFecha.ToString("yyyyMMdd"), iBodega);
					sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL, true);

					if (sMensaje.Length > 0)
					{
						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

						this.txtSerial.Text = "";
						this.txtIdArticulo.Value = 0;
						this.txtCodigo.Text = "";
						this.txtSerial.Focus();
						return;
					}
				
					sMensaje = "";
					string sSQLSerial = string.Format("Exec ValidaSerialExistenteEnSeriales {0}, '{1}'", (int)this.txtIdArticulo.Value, this.txtSerial.Text.ToString());
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLSerial, true);
					dr.Read();
					if (dr.HasRows)
					{
						iRes = dr.GetInt32(0);
						sMensaje = dr.GetString(1);
						idSerial = dr.GetInt32(2);
					}
					dr.Close();
				
					if (iRes == 1) // SERIAL NO EXISTE
					{
						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtSerial.Text = "";
						this.txtIdArticulo.Value = 0;
						this.txtCodigo.Text = "";
						this.txtSerial.Focus();
						return;
					}

					if (iRes == 2) // SERIAL NO EXISTE PERO REGISTRA CON EL ARTICULO SELECCIONADO
					{
						if (DialogResult.Yes == MessageBox.Show(sMensaje + " " + this.txtCodigo.Text.ToString(), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
						{
							string sSQLGC = string.Format("Exec GuardaConteoInventario {0}, '{1}', {2}, {3}, '{4}', {5}, {6}, {7},{8}",
								(int)this.txtIdArticulo.Value, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, idSerial, this.txtSerial.Text.ToString(), 1, 1, 0,(bool) this.optBodegaExhibicion.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLGC, true);

							this.btnVerToma_Click(sender, e);

							this.txtSerial.Text = "";
							this.txtIdArticulo.Value = 0;
							this.txtCodigo.Text = "";

							return;
						}
					}

					if (iRes == 3) // SERIAL SI EXISTE CAMBIA ARTICULO ACTUAL DE LA SERIAL POR EL SELECCIONADO 
					{
						if (DialogResult.Yes == MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
						{
							string sSQLGC = string.Format("Exec GuardaConteoInventario {0}, '{1}', {2}, {3}, '{4}', {5}, {6}, {7},{8}",
								(int)this.txtIdArticulo.Value, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, idSerial, this.txtSerial.Text.ToString(), 1, 1, 1,(bool) this.optBodegaExhibicion.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLGC, true);

							this.btnVerToma_Click(sender, e);

							this.txtSerial.Text = "";
							this.txtIdArticulo.Value = 0;
							this.txtCodigo.Text = "";

							return;
						}
					}

					if (iRes == 4) // SERIAL SI EXISTE MANTIENE ARTICULO ACTUAL
					{
						string sSQLGC = string.Format("Exec GuardaConteoInventario {0}, '{1}', {2}, {3}, '{4}', {5}, {6}, {7},{8}",
							(int)this.txtIdArticulo.Value, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, idSerial, this.txtSerial.Text.ToString(), 1, 1, 0,(bool) this.optBodegaExhibicion.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLGC, true);

						this.btnVerToma_Click(sender, e);

						this.txtSerial.Text = "";
						this.txtIdArticulo.Value = 0;
						this.txtCodigo.Text = "";

						return;						
					}
					#endregion Maneja Serial
				}
			}
			#endregion Valida Con Articulo Seleccionado

			#region Valida Sin Articulo Seleccionado
			if ((int)this.txtIdArticulo.Value == 0)
			{				
				string sSQL = string.Format("Exec ValidaSerialEnConteo {0}, '{1}', '{2}', {3}", (int)this.txtIdArticulo.Value, this.txtSerial.Text.ToString(), dtFecha.ToString("yyyyMMdd"), iBodega);
				sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL, true);

				if (sMensaje.Length > 0)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerial.Focus();
					return;
				}
				
				sMensaje = "";
				string sSQLSerial = string.Format("Exec ValidaSerialExistenteEnSeriales {0}, '{1}'", (int)this.txtIdArticulo.Value, this.txtSerial.Text.ToString());
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLSerial, true);
				dr.Read();
				if (dr.HasRows)
				{
					iRes = dr.GetInt32(0);
					sMensaje = dr.GetString(1);
					idSerial = dr.GetInt32(2);
					idArticulo = dr.GetInt32(3);
				}
				dr.Close();
				
				if (iRes == 1)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerial.Focus();
					return;
				}

				if (iRes == 4)
				{
					string sSQLGC = string.Format("Exec GuardaConteoInventario {0}, '{1}', {2}, {3}, '{4}', {5}, {6}, {7},{8}",
						idArticulo, dtFecha.ToString("yyyyMMdd HH:mm"), iBodega, idSerial, this.txtSerial.Text.ToString(), 1, 1, 0,(bool) this.optBodegaExhibicion.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLGC, true);

					this.btnVerToma_Click(sender, e);

					this.txtSerial.Text = "";
					this.txtIdArticulo.Value = 0;
					this.txtCodigo.Text = "";

					return;					
				}				
			}
			#endregion Valida Sin Articulo Seleccionado

			MsjeContExis();
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			this.txtFecha.Value = DateTime.Now;
		}

		private void txtBuscaSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				DateTime Fecha = (DateTime)this.txtFecha.Value;
				string sSQL = string.Format("Exec ConteoBuscaSerial '{0}', {1}, '{2}'", Fecha.ToString("yyyyMMdd"), iBodega, this.txtBuscaSerial.Text);
				this.grdDTomas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	
			}
		}

		private void frmTomaInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnVerToma_Click(object sender, System.EventArgs e)
		{
			Cursor.Current = Cursors.WaitCursor;						
			int NumeroToma = 1; 
			
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size=75000;"))
			{
				oConexion.Open();

				SqlCommand oCmdToma = oConexion.CreateCommand();
				oCmdToma.Connection = oConexion;				
				oCmdToma.CommandTimeout = 0;
				oCmdToma.CommandType = CommandType.StoredProcedure;
				oCmdToma.CommandText = "Visualiza_TomaInventario";

				oCmdToma.Parameters.Add("@NToma", SqlDbType.Int).Value = NumeroToma;
				oCmdToma.Parameters.Add("@Bodega", SqlDbType.Int).Value = iBodega;
				oCmdToma.Parameters.Add("@FechToma", SqlDbType.DateTime).Value = (DateTime)this.txtFecha.Value;;
				oCmdToma.Parameters.Add("@CodArti", SqlDbType.VarChar).Value = this.txtCodigo.Text.ToString().Trim();

				SqlDataAdapter oDAToma = new SqlDataAdapter(oCmdToma);
				DataTable oDTToma = new DataTable();
				oDAToma.Fill(oDTToma);
				this.grdDTomas.DataSource = oDTToma;

				SqlCommand oCmdConteo = oConexion.CreateCommand();
				oCmdConteo.Connection = oConexion;				
				oCmdConteo.CommandTimeout = 0;			
				oCmdConteo.CommandType = CommandType.StoredProcedure;
				oCmdConteo.CommandText = "Visualiza_Conteo_TomaInventario";
				
				oCmdConteo.Parameters.Add("@NToma", SqlDbType.Int).Value = NumeroToma;
				oCmdConteo.Parameters.Add("@Bodega", SqlDbType.Int).Value = iBodega;
				oCmdConteo.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = (DateTime)this.txtFecha.Value;;				

				SqlDataAdapter oDAConteo = new SqlDataAdapter(oCmdConteo);
				DataTable oDTConteo = new DataTable();
				oDAConteo.Fill(oDTConteo);
				this.grContoInv.DataSource = oDTConteo;
			
				oConexion.Close();
			}		

			MsjeContExis();

			Cursor.Current = Cursors.Default;		
		}

		private void txtBuscaSerial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBuscaSerial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSerial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
