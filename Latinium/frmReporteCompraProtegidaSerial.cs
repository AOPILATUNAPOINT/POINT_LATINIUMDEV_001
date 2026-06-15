using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.Globalization;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteCompraProtegidaSerial.
	/// </summary>
	public class frmReporteCompraProtegidaSerial : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TextBox txtSerial;
		private System.Windows.Forms.Label label72;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdSerial;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		int idArticulo = 0;
		int idCompra = 0;
		int idCompraOrigen = 0;
		public static int idSerial = 0;
		bool ModSerial = false;
		private System.Windows.Forms.TextBox txtArticulo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtBodega;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtCodigoArticulo;
		private System.Windows.Forms.Label label3;

		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteCompraProtegidaSerial(int IdCompra)
		{
			idCompra = IdCompra;
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
		/// 
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteCompraProtegidaSerial));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Consignacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.txtSerial = new System.Windows.Forms.TextBox();
			this.label72 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.grdSerial = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtArticulo = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.txtBodega = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtCodigoArticulo = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.grdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			this.txtSerial.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.Location = new System.Drawing.Point(104, 112);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(304, 20);
			this.txtSerial.TabIndex = 17;
			this.txtSerial.Text = "";
			this.txtSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerial_KeyDown);
			this.txtSerial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerial_KeyPress);
			// 
			// label72
			// 
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Location = new System.Drawing.Point(8, 112);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(112, 16);
			this.label72.TabIndex = 154;
			this.label72.Text = "* Ingrese la Serial:";
			this.label72.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 152);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 40;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(104, 152);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 155;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// grdSerial
			// 
			this.grdSerial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.grdSerial.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdSerial.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 28;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 41;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 85;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 253;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 446;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 34;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 55;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.grdSerial.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdSerial.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.grdSerial.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdSerial.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.grdSerial.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 9F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSerial.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdSerial.Enabled = false;
			this.grdSerial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSerial.Location = new System.Drawing.Point(280, 152);
			this.grdSerial.Name = "grdSerial";
			this.grdSerial.Size = new System.Drawing.Size(48, 24);
			this.grdSerial.TabIndex = 156;
			this.grdSerial.Visible = false;
			this.grdSerial.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grdSerial_InitializeRow);
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(224, 160);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(43, 16);
			this.label15.TabIndex = 426;
			this.label15.Text = "Bodega";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Visible = false;
			// 
			// cmbBodega
			// 
			this.cmbBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 180;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(192, 152);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(24, 22);
			this.cmbBodega.TabIndex = 425;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
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
			// txtArticulo
			// 
			this.txtArticulo.AcceptsTab = true;
			this.txtArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.Location = new System.Drawing.Point(104, 48);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.ReadOnly = true;
			this.txtArticulo.Size = new System.Drawing.Size(304, 20);
			this.txtArticulo.TabIndex = 427;
			this.txtArticulo.Text = "";
			// 
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(112, 16);
			this.label1.TabIndex = 428;
			this.label1.Text = "Articulo:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBodega
			// 
			this.txtBodega.AcceptsTab = true;
			this.txtBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBodega.Location = new System.Drawing.Point(104, 16);
			this.txtBodega.Name = "txtBodega";
			this.txtBodega.ReadOnly = true;
			this.txtBodega.Size = new System.Drawing.Size(304, 20);
			this.txtBodega.TabIndex = 429;
			this.txtBodega.Text = "";
			// 
			// label2
			// 
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(112, 16);
			this.label2.TabIndex = 430;
			this.label2.Text = "Bodega:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigoArticulo
			// 
			this.txtCodigoArticulo.AcceptsTab = true;
			this.txtCodigoArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtCodigoArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigoArticulo.Location = new System.Drawing.Point(104, 80);
			this.txtCodigoArticulo.Name = "txtCodigoArticulo";
			this.txtCodigoArticulo.ReadOnly = true;
			this.txtCodigoArticulo.Size = new System.Drawing.Size(304, 20);
			this.txtCodigoArticulo.TabIndex = 431;
			this.txtCodigoArticulo.Text = "";
			// 
			// label3
			// 
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(112, 16);
			this.label3.TabIndex = 432;
			this.label3.Text = "Código artículo:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmReporteCompraProtegidaSerial
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(424, 182);
			this.Controls.Add(this.txtCodigoArticulo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtBodega);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.grdSerial);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.label72);
			this.Name = "frmReporteCompraProtegidaSerial";
			this.Text = "Ingreso de serial";
			this.Load += new System.EventHandler(this.frmReporteCompraProtegidaSerial_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteCompraProtegidaSerial_Load(object sender, System.EventArgs e)
		{
			idSerial = 0;
			idArticulo = 0;
			string squery = string.Format("Exec ConsultaDetalleFactura {0}", idCompra);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, squery, true);
			dr.Read();
			if (dr.HasRows)
			{
				txtCodigoArticulo.Text = dr.GetString(5);
				txtArticulo.Text = dr.GetString(6);
				idArticulo = dr.GetInt32(3);
			}
			dr.Close();
			//La bodega hace el mach a la bodega de reclamos seguros 151 en la tabla Bodega
			this.cmbBodega.Value = 151;
			//Obtener descripción de la bodega RECLAMOS SEGUROS
			this.txtBodega.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("Select Nombre From Bodega where Bodega = {0}", this.cmbBodega.Value));

		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if(this.txtSerial.Text == "")
			{
				MessageBox.Show("El campo serial es obligatorio.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtSerial.Focus();
				return;
			}
			ValidaSerial();
//			if(idSerial == 0)
//			{
//				MessageBox.Show("No se a ingresado un serial correcto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//				this.txtSerial.Focus();
//				return;
//			}
			if(idSerial != 0 && this.txtSerial.Text != "")
			{
				this.DialogResult = DialogResult.OK;
			}
		}

		private void txtSerial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
				if (Char.IsControl(e.KeyChar))
				{
					e.Handled = false;
				}
				else if (Char.IsLetterOrDigit(e.KeyChar))
				{
					e.Handled = false;
				}			
				else
				{
					e.Handled = true;
				}
		}

		private void txtSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
			{
				ValidaSerial();
			}
		}

		private void ValidaSerial()
		{
			idSerial = 0;
			int IdDetalleSerial = 0;
			bool bConsignacion = false;
			bool SerialRepetida = false;
			int CantSerie = 0;

			#region Cuenta si hay articulos en el grid
			if (idArticulo == 0)
			{
				MessageBox.Show("Ingrese el artículo para cargar seriales.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtSerial.Focus();
				return;
			}
			#endregion Cuenta si hay articulos en el grid

			#region Valida Selección 
			if (idArticulo == 0)
			{
				MessageBox.Show("Seleccione el Artículo", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtSerial.Text = "";
				this.txtSerial.Focus();
				return;
			}
			#endregion Valida Selección 
			int IdArticulo = idArticulo;
			int IPosicion = 0; 
			#region Valida cantidad seriales
			if (SerialRepetida == false)
			{			      
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.grdSerial.Rows.All)
				{
					int idArtSerial = (int)dr1.Cells["idArticulo"].Value;
					int iPosSerial = 0;

					if (IdArticulo == idArtSerial && IPosicion == iPosSerial)
					{
						CantSerie++;
					}
				}
			}
			#endregion Valida cantidad seriales
			#region Valida serial existente en el grid
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdSerial.Rows.All)
			{
				string sSerial = dr.Cells["Serial"].Value.ToString();
				if ((string)this.txtSerial.Text == sSerial)
				{
					MessageBox.Show("Ya registro esta serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}
			}
			#endregion Valida serial existente en el grid
			#region Valida si la Serial Existe en el Sistema o no Pertenece al Articulo Seleccionado
			string sSQLVSA = string.Format("Exec ValidaSerialArticulo '{0}', {1}", this.txtSerial.Text.ToString(), IdArticulo);
			string sMensajeVSA = Funcion.sEscalarSQL(cdsSeteoF, sSQLVSA);

			if (sMensajeVSA.Length > 0)
			{
				MessageBox.Show(sMensajeVSA, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerial.Text = "";
				this.txtSerial.Focus();
				return;
			}
			#endregion Valida si la Serial Existe en el Sistema o no Pertenece al Articulo Seleccionado		
			#region Valida si la Serial esta en Orden de Compra - En Transito - Pedidos - Reservas - Consignaciones - Ventas 
			int iCont = 0;
			string sMens = "";
			string sSQLValSerEsp = string.Format("Exec ValidaSerialOrdCmpEnTransito '{0}', {1}, {2}, 1, {3}", this.txtSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega.Value, idCompraOrigen);
			SqlDataReader drV = Funcion.rEscalarSQL(cdsSeteoF, sSQLValSerEsp, true);
			drV.Read();
			if (drV.HasRows)
			{
				iCont = drV.GetInt32(0);
				sMens = drV.GetString(1);
			}
			drV.Close();
								
			if (iCont > 0)
			{
				MessageBox.Show(sMens, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerial.Focus();
				return;
			}
			#endregion Valida si la Serial esta en Orden de Compra - En Transito - Pedidos - Reservas - Consignaciones - Ventas
			#region Carga Grid de Seriales
			string sSQL = string.Format("Exec CuentaSerialArticulo '{0}', {1}, {2}, {3}", this.txtSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega.Value, idCompraOrigen);
			int iCtaSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
        
			if (iCtaSerial == 1)
			{	
				idSerial = 0;
				IdDetalleSerial = 0;

				string sSQLidSerial = string.Format("Exec BuscaidSerial '{0}', {1}, {2}, {3}", this.txtSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega.Value, idCompraOrigen);
				SqlDataReader dridSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLidSerial, true);
				dridSerial.Read();
				if (dridSerial.HasRows)
				{
					idSerial = dridSerial.GetInt32(0);
					IdDetalleSerial = dridSerial.GetInt32(1);
					bConsignacion = dridSerial.GetBoolean(2);
				}
				dridSerial.Close();
					
				//					if (!ModSerial) 
				//					{																		
				//						this.grdSerial.Rows.Band.AddNew();
				//						this.grdSerial.ActiveRow.Cells["idSerial"].Value = idSerial;
				//						this.grdSerial.ActiveRow.Cells["idDetalleSerial"].Value = 0;
				//						this.grdSerial.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
				//						this.grdSerial.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
				//						this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();	
				//						this.grdSerial.ActiveRow.Cells["Estado"].Value = 4;	
				//						this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
				//						this.grdSerial.ActiveRow.Cells["Consignacion"].Value = bConsignacion;
				//					}
				//					else
				//					{
				//						this.grdSerial.ActiveRow.Cells["idSerial"].Value = idSerial;						
				//						this.grdSerial.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;						
				//						this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtSerial.Text.ToString();	
				//						this.grdSerial.ActiveRow.Cells["Estado"].Value = 4;	
				//						this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;
				//						this.grdSerial.ActiveRow.Cells["Consignacion"].Value = bConsignacion;
				//					}
			}
			else
			{
				MessageBox.Show("La serial no existe, Esta en otro Local o no pertenece al artículo seleccionado,\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtSerial.Focus();
				return;
			}

			MessageBox.Show("La serial fue validada correctamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			this.txtSerial.Focus();
			
			//this.txtSerial.Text = "";
			#endregion Carga Grid de Seriales
			this.txtSerial.Focus();
			ModSerial = false;
		}

		private void grdSerial_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if ((int)e.Row.Cells["Estado"].Value == 11)
			{
				e.Row.Appearance.BackColor = Color.Red;
				e.Row.Appearance.BackColor2 = Color.Red;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}
	}
}
