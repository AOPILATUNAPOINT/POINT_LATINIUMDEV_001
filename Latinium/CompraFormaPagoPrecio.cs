using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Windows.Forms;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraFormaPago.
	/// </summary>
	public class CompraFormaPagoPrecio : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor IdFormaPago;
		private System.Windows.Forms.Label label13;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor idCliente;
		public System.Windows.Forms.TextBox txtDocumento;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor Cliente;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor Saldo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor RUC;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEfectivo;

		int IdTipoFactura = 0;
		public string sNumero = "";
		public int idTipoRuc = 0;
		public int IdCompraNC = 0;
		public bool NotaCredito = false;
		private C1.Data.C1DataSet cdsSeteoF;

		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		
		public CompraFormaPagoPrecio(int idTipoFactura)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			IdTipoFactura = idTipoFactura;
			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.IdFormaPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.idCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDocumento = new System.Windows.Forms.TextBox();
			this.Cliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.Saldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.RUC = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.IdFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Cliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.Saldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.RUC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
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
			ultraGridBand1.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 130;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 283;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(1, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(304, 312);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(184, 320);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(112, 25);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// IdFormaPago
			// 
			this.IdFormaPago.Location = new System.Drawing.Point(408, 352);
			this.IdFormaPago.Name = "IdFormaPago";
			this.IdFormaPago.Size = new System.Drawing.Size(64, 22);
			this.IdFormaPago.TabIndex = 5;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(8, 320);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(61, 16);
			this.label13.TabIndex = 133;
			this.label13.Text = "Documento";
			this.label13.Visible = false;
			// 
			// idCliente
			// 
			this.idCliente.Location = new System.Drawing.Point(408, 376);
			this.idCliente.Name = "idCliente";
			this.idCliente.Size = new System.Drawing.Size(64, 22);
			this.idCliente.TabIndex = 135;
			this.idCliente.Text = "0";
			// 
			// txtDocumento
			// 
			this.txtDocumento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDocumento.Location = new System.Drawing.Point(72, 320);
			this.txtDocumento.Name = "txtDocumento";
			this.txtDocumento.Size = new System.Drawing.Size(104, 20);
			this.txtDocumento.TabIndex = 137;
			this.txtDocumento.Text = "";
			this.txtDocumento.Visible = false;
			this.txtDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDocumento_KeyDown);
			this.txtDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtDocumento_Validating);
			// 
			// Cliente
			// 
			this.Cliente.Location = new System.Drawing.Point(408, 328);
			this.Cliente.Name = "Cliente";
			this.Cliente.Size = new System.Drawing.Size(64, 22);
			this.Cliente.TabIndex = 138;
			// 
			// Saldo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.Saldo.Appearance = appearance7;
			this.Saldo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.Saldo.FormatString = "#,##0.00";
			this.Saldo.Location = new System.Drawing.Point(368, 304);
			this.Saldo.MinValue = 0;
			this.Saldo.Name = "Saldo";
			this.Saldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.Saldo.PromptChar = ' ';
			this.Saldo.Size = new System.Drawing.Size(104, 22);
			this.Saldo.TabIndex = 139;
			// 
			// RUC
			// 
			this.RUC.Location = new System.Drawing.Point(408, 280);
			this.RUC.Name = "RUC";
			this.RUC.Size = new System.Drawing.Size(64, 22);
			this.RUC.TabIndex = 140;
			// 
			// txtEfectivo
			// 
			this.txtEfectivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEfectivo.Appearance = appearance8;
			this.txtEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEfectivo.FormatString = "#,##0.00";
			this.txtEfectivo.Location = new System.Drawing.Point(72, 320);
			this.txtEfectivo.Name = "txtEfectivo";
			this.txtEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEfectivo.PromptChar = ' ';
			this.txtEfectivo.Size = new System.Drawing.Size(104, 22);
			this.txtEfectivo.TabIndex = 141;
			this.txtEfectivo.Visible = false;
			this.txtEfectivo.Click += new System.EventHandler(this.txtEfectivo_Click);
			this.txtEfectivo.Validating += new System.ComponentModel.CancelEventHandler(this.txtEfectivo_Validating);
			this.txtEfectivo.ValueChanged += new System.EventHandler(this.txtEfectivo_ValueChanged);
			this.txtEfectivo.Enter += new System.EventHandler(this.txtEfectivo_Enter);
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
			// CompraFormaPagoPrecio
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(306, 344);
			this.Controls.Add(this.txtEfectivo);
			this.Controls.Add(this.RUC);
			this.Controls.Add(this.Saldo);
			this.Controls.Add(this.Cliente);
			this.Controls.Add(this.txtDocumento);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.idCliente);
			this.Controls.Add(this.IdFormaPago);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "CompraFormaPagoPrecio";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione la forma de pago";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.CompraFormaPagoPrecio_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CompraFormaPagoPrecio_Closing);
			this.Load += new System.EventHandler(this.CompraFormaPagoPrecio_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.IdFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Cliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.Saldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.RUC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

//		public int idCliente;
//		public double Saldo;		

		public void Aceptar()
		{
			if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 4)
			{
				if (this.txtDocumento.TextLength > 0)
				{
					string stSelect = string.Format("Exec VentaDevolucion '{0}'", txtDocumento.Text);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
                
					dr.Read();
					if (dr.HasRows)
					{
						idCliente.Value = (int)dr.GetValue(0);
						RUC.Value = (string)dr.GetValue(1);
						Cliente.Value = (string)dr.GetValue(2);
						IdCompraNC = (int)dr.GetValue(3);
						Saldo.Value = (double)dr.GetValue(4);
						IdFormaPago.Value = (int)dr.GetValue(5);
						sNumero = (string)dr.GetValue(6);
						idTipoRuc = (int)dr.GetValue(7);
						NotaCredito = true;
						dr.Close();
						this.DialogResult = DialogResult.OK;
					}
					else
					{
						dr.Close();
						MessageBox.Show("El documento no existe o no esta autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						txtDocumento.Focus();
						return;
					}
				}
				else
				{
					MessageBox.Show("Ingrese el número de documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					txtDocumento.Focus();
					return;
				}
			}
			else if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 5 && (double)this.txtEfectivo.Value == 0)
			{
				MessageBox.Show("Ingrese el valor de efectivo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			else
			{
				this.IdFormaPago.Value = (int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;

				this.DialogResult = DialogResult.OK;
			}
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();			
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void CompraFormaPagoPrecio_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			this.IdFormaPago.Width = 0;
			this.idCliente.Width = 0;
			this.RUC.Width = 0;
			this.Saldo.Width = 0;
			this.Cliente.Width = 0;
			this.idCliente.Value = 0;
			this.Saldo.Value = 0;

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FormasDePagoVentas {0}, {1}", IdTipoFactura, MenuLatinium.iNivel));
			Cursor = Cursors.Default;
		}

		private void CompraFormaPagoPrecio_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{

		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			this.label13.Visible = true;
			this.txtDocumento.Visible = true;

			this.IdFormaPago.Value = (int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;

			if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 4)
			{
				this.label13.Text = "Documento";
				this.label13.Visible = true;
				this.txtDocumento.Visible = true;
				this.txtEfectivo.Visible = false;										
				this.txtDocumento.Focus();
			}
			else if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 5)
			{
				this.label13.Text = "Efectivo";
				this.label13.Visible = true;
				this.txtEfectivo.Visible = true;
				this.txtDocumento.Visible = false;
				this.txtEfectivo.Focus();
			}
			else
			{
				this.label13.Visible = false;
				this.txtDocumento.Visible = false;
				this.txtEfectivo.Visible = false;			
			}
		}

		private void txtDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}
		
		private void txtDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void txtEfectivo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtEfectivo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void txtEfectivo_Click(object sender, System.EventArgs e)
		{
			this.txtEfectivo.SelectAll();
		}

		private void txtEfectivo_Enter(object sender, System.EventArgs e)
		{
			this.txtEfectivo.SelectAll();
		}

		private void CompraFormaPagoPrecio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				this.DialogResult = DialogResult.No;
				Close();
			}
		}

		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) Aceptar();
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{			
			this.label13.Visible = true;
			this.txtDocumento.Visible = true;

			this.IdFormaPago.Value = (int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;

			if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 4)
			{
				this.label13.Text = "Documento";
				this.label13.Visible = true;
				this.txtDocumento.Visible = true;
				this.txtEfectivo.Visible = false;										
				this.txtDocumento.Focus();
			}
			else if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 5)
			{
				this.label13.Text = "Efectivo";
				this.label13.Visible = true;
				this.txtEfectivo.Visible = true;
				this.txtDocumento.Visible = false;
				this.txtEfectivo.Focus();
			}
			else
			{
				this.label13.Visible = false;
				this.txtDocumento.Visible = false;
				this.txtEfectivo.Visible = false;			
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}	
	}
}
