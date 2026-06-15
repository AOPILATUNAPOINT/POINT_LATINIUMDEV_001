using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPrefacturaBonoEntradaSel.
	/// </summary>
	public class frmPrefacturaBonoEntradaSel : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsBonoEntrada;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnAceptar;
		/// <summary>
		/// Variable del diseñador requerida.
		public int idBonoEntrada = 0;
		public int idCliente = 0;
		public int idBonoEntradaTipo = 0;

		private Infragistics.Win.UltraWinGrid.UltraGrid udgBonoEntrada;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPrefacturaBonoEntradaSel(int IdCliente, int IdBonoEntradaTipo)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idCliente = IdCliente;
			idBonoEntradaTipo = IdBonoEntradaTipo;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBonoEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBonoEntradaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBonoEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBonoEntradaTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEmision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			this.udgBonoEntrada = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsBonoEntrada = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAceptar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.udgBonoEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBonoEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// udgBonoEntrada
			// 
			this.udgBonoEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgBonoEntrada.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgBonoEntrada.DataSource = this.udsBonoEntrada;
			appearance1.BackColor = System.Drawing.Color.White;
			this.udgBonoEntrada.DisplayLayout.Appearance = appearance1;
			this.udgBonoEntrada.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 24;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 33;
			ultraGridColumn3.AutoEdit = false;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Format = "dd/MM/yyyy";
			ultraGridColumn3.Header.Caption = "Fecha emisión";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Width = 117;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 19;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 26;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 34;
			ultraGridColumn7.AutoEdit = false;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Número";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 118;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.Caption = "Estado";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.AutoEdit = false;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "N°";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 73;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand1.Header.Caption = "Facturas del cliente";
			ultraGridBand1.HeaderVisible = true;
			this.udgBonoEntrada.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.udgBonoEntrada.DisplayLayout.GroupByBox.Hidden = true;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgBonoEntrada.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.udgBonoEntrada.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgBonoEntrada.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgBonoEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.udgBonoEntrada.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgBonoEntrada.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.udgBonoEntrada.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgBonoEntrada.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgBonoEntrada.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgBonoEntrada.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.udgBonoEntrada.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgBonoEntrada.FlatMode = true;
			this.udgBonoEntrada.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgBonoEntrada.Location = new System.Drawing.Point(16, 16);
			this.udgBonoEntrada.Name = "udgBonoEntrada";
			this.udgBonoEntrada.Size = new System.Drawing.Size(256, 200);
			this.udgBonoEntrada.TabIndex = 763;
			this.udgBonoEntrada.Click += new System.EventHandler(this.udgBonoEntrada_Click);
			// 
			// udsBonoEntrada
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.udsBonoEntrada.Band.Columns.AddRange(new object[] {
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
			// 
			// btnAceptar
			// 
			this.btnAceptar.Location = new System.Drawing.Point(16, 232);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 764;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// frmPrefacturaBonoEntradaSel
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 266);
			this.ControlBox = false;
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.udgBonoEntrada);
			this.Name = "frmPrefacturaBonoEntradaSel";
			this.Text = "Seleccionar Bono";
			this.Load += new System.EventHandler(this.frmPrefacturaBonoEntradaSel_Load);
			((System.ComponentModel.ISupportInitialize)(this.udgBonoEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsBonoEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPrefacturaBonoEntradaSel_Load(object sender, System.EventArgs e)
		{

			string squery = String.Format("Exec BonoEntradaClienteLista {0}, {1}",
				idCliente,
				idBonoEntradaTipo);

			FuncionesProcedimientos.dsGeneralModulo(squery, this.udgBonoEntrada);

		}

		protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
		{
			// Verificar si la tecla presionada es la tecla Escape
			if (keyData == Keys.Escape)
			{
				// Devolver true para indicar que el evento de tecla ha sido manejado
				return true;
			}

			// Verificar si las teclas presionadas son Alt + F4
			if (keyData == (Keys.Alt | Keys.F4))
			{
				// Devolver true para indicar que el evento de tecla ha sido manejado
				return true;
			}

			// Llamar al método de la clase base para otras teclas
			return base.ProcessCmdKey(ref msg, keyData);
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (idBonoEntrada == 0)
			{
				MessageBox.Show("Seleccione una factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			this.DialogResult = DialogResult.OK;
		}

		private void udgBonoEntrada_Click(object sender, System.EventArgs e)
		{
		 idBonoEntrada = (int)this.udgBonoEntrada.ActiveRow.Cells["idBonoEntrada"].Value;
		}
	}
}
