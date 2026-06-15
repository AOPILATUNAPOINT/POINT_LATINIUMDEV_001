using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for P_BuscarOrdenPago.
	/// </summary>
	public class P_BuscarOrdenPago : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private DevExpress.XtraEditors.SimpleButton BtBusca;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private DevExpress.XtraEditors.SimpleButton btCancela;
		private DevExpress.XtraEditors.SimpleButton btAcepta;
		private C1.Data.C1DataSet cdsSolicitudPago;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public P_BuscarOrdenPago()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.BtBusca = new DevExpress.XtraEditors.SimpleButton();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btCancela = new DevExpress.XtraEditors.SimpleButton();
			this.btAcepta = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSolicitudPago = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSolicitudPago)).BeginInit();
			this.SuspendLayout();
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(-118, -8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(57, 25);
			this.btBuscar.TabIndex = 12;
			this.btBuscar.Text = "&Buscar";
			// 
			// txtArticulo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance1;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(-230, -8);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(101, 22);
			this.txtArticulo.TabIndex = 11;
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(622, 223);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 15;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(-190, 223);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(76, 24);
			this.btAceptar.TabIndex = 14;
			this.btAceptar.Text = "&Aceptar";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(496, 152);
			this.ultraGrid1.TabIndex = 18;
			// 
			// BtBusca
			// 
			this.BtBusca.Location = new System.Drawing.Point(128, 24);
			this.BtBusca.Name = "BtBusca";
			this.BtBusca.Size = new System.Drawing.Size(56, 25);
			this.BtBusca.TabIndex = 17;
			this.BtBusca.Text = "&Buscar";
			this.BtBusca.Click += new System.EventHandler(this.BtBusca_Click);
			// 
			// txtBuscar
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance7;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(16, 24);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(100, 22);
			this.txtBuscar.TabIndex = 16;
			// 
			// btCancela
			// 
			this.btCancela.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancela.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancela.Location = new System.Drawing.Point(408, 208);
			this.btCancela.Name = "btCancela";
			this.btCancela.Size = new System.Drawing.Size(75, 24);
			this.btCancela.TabIndex = 20;
			this.btCancela.Text = "&Cancelar";
			// 
			// btAcepta
			// 
			this.btAcepta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAcepta.Location = new System.Drawing.Point(48, 208);
			this.btAcepta.Name = "btAcepta";
			this.btAcepta.Size = new System.Drawing.Size(75, 24);
			this.btAcepta.TabIndex = 19;
			this.btAcepta.Text = "&Aceptar";
			this.btAcepta.Click += new System.EventHandler(this.btAcepta_Click);
			// 
			// cdsSolicitudPago
			// 
			this.cdsSolicitudPago.BindingContextCtrl = this;
			this.cdsSolicitudPago.DataLibrary = "LibFacturacion";
			this.cdsSolicitudPago.DataLibraryUrl = "";
			this.cdsSolicitudPago.DataSetDef = "dsPresupuestoC";
			this.cdsSolicitudPago.FillOnRequest = false;
			this.cdsSolicitudPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSolicitudPago.Name = "cdsSolicitudPago";
			this.cdsSolicitudPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSolicitudPago.SchemaDef = null;
			this.cdsSolicitudPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsSolicitudPago_BeforeFill);
			// 
			// P_BuscarOrdenPago
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(520, 238);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.BtBusca);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btCancela);
			this.Controls.Add(this.btAcepta);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Name = "P_BuscarOrdenPago";
			this.Text = "Buscar Solicitud de Pago";
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSolicitudPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		public int idSolicitudPago =0;
		private void BtBusca_Click(object sender, System.EventArgs e)
		{			
			string stSelect = string.Format("Exec P_PresupuestoBuscaSolicitudP '{0}'", txtBuscar.Value);
			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSolicitudPago, stSelect);
			ultraGrid1.DisplayLayout.Bands[0].Columns["idSolicitudPago"].Hidden = true;

			if (ultraGrid1.Rows.Count == 1)
			{
				ultraGrid1.Select();
				if (ultraGrid1.ActiveRow == null ||	ultraGrid1.ActiveRow.Cells["idSolicitudPago"].Value == DBNull.Value)
					return;
				idSolicitudPago = (int) ultraGrid1.ActiveRow.Cells["idSolicitudPago"].Value;

				DialogResult = DialogResult.OK;
			}
			else
			{
				AcceptButton=btAceptar;
				ultraGrid1.Select();
			}
		}

		private void cdsSolicitudPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSolicitudPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAcepta_Click(object sender, System.EventArgs e)
		{
			try
			{
				idSolicitudPago = (int) ultraGrid1.ActiveRow.Cells["idSolicitudPago"].Value;
				DialogResult = DialogResult.OK;
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe seleccionar un registro..."+ex.Message,"Error");
			}
		}
	}
}

