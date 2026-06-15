using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmFormasDePago.
	/// </summary>
	public class frmFormasDePago : DevExpress.XtraEditors.XtraForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label13;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumNotaCredito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;

		int IdTipoFactura = 0;
		bool bFDPCredito = false;
		bool bFDPContraEntrega = false;
		int iBodega = 0;

		public frmFormasDePago(int idTipoFactura, bool BFDPCredito, bool BFDPContraEntrega, int IBodega)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			IdTipoFactura = idTipoFactura;
			bFDPCredito = BFDPCredito;
			bFDPContraEntrega = BFDPContraEntrega;
			iBodega = IBodega;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promocional");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promocional");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtNumNotaCredito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumNotaCredito)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 130;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Forma de Pago";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 299;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 63;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
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
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(320, 320);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(8, 339);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(141, 16);
			this.label13.TabIndex = 142;
			this.label13.Text = "Número de Nota de Crédito";
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
			// txtNumNotaCredito
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumNotaCredito.Appearance = appearance7;
			this.txtNumNotaCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumNotaCredito.Enabled = false;
			this.txtNumNotaCredito.Location = new System.Drawing.Point(160, 336);
			this.txtNumNotaCredito.Name = "txtNumNotaCredito";
			this.txtNumNotaCredito.Size = new System.Drawing.Size(152, 22);
			this.txtNumNotaCredito.TabIndex = 1;
			this.txtNumNotaCredito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumNotaCredito_KeyPress);
			this.txtNumNotaCredito.Validated += new System.EventHandler(this.txtNumNotaCredito_Validated);
			this.txtNumNotaCredito.ValueChanged += new System.EventHandler(this.txtNumNotaCredito_ValueChanged);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 368);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 2;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(88, 368);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 3;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// frmFormasDePago
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(320, 398);
			this.ControlBox = false;
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtNumNotaCredito);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmFormasDePago";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione la Forma de pago";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFormasDePago_KeyDown);
			this.Load += new System.EventHandler(this.frmFormasDePago_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumNotaCredito)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public int idCliente = 0;
		public int idTipoRuc = 0;
		public string RUC = "";
		public string Cliente = "";
		public int IdCompraNC = 0;
		public string sNumero = "";
		public int IdFormaPago = 0;
		public decimal dSaldo = 0.00m;
		public bool NotaCredito = false;
		public int IdFormaPagoNC = 0;
		
		public void Aceptar()
		{
			if ((int) this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 4)
			{
				if (this.txtNumNotaCredito.TextLength > 0)
				{
					string stSelect = string.Format("Exec VentaDevolucion '{0}'", this.txtNumNotaCredito.Text.ToString());
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
                
					dr.Read();
					if (dr.HasRows)
					{
						idCliente = (int)dr.GetValue(0);
						RUC = (string)dr.GetValue(1);
						Cliente = (string)dr.GetValue(2);
						IdCompraNC = (int)dr.GetValue(3);
						dSaldo = dr.GetDecimal(4);
						IdFormaPagoNC = (int)dr.GetValue(5);
						sNumero = (string)dr.GetValue(6);
						idTipoRuc = (int)dr.GetValue(7);
						NotaCredito = true;
						dr.Close();

						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From BloqueoTransacciones Where idOrigen = {0} And Origen = 1 And Bloqueado = 1", IdCompraNC)) > 0)
						{
							string sUsuario = "";
							string sEstacion = "";
							DateTime dtFechaBloqueo = DateTime.Today;
						
							string sSQLValB = string.Format("Select Usuario, Estacion, Fecha From BloqueoTransacciones Where idOrigen = {0} And Origen = 1", IdCompraNC);
							SqlDataReader drBloqueo = Funcion.rEscalarSQL(cdsSeteoF, sSQLValB, true);
							drBloqueo.Read();
							if (drBloqueo.HasRows)
							{
								sUsuario = drBloqueo.GetValue(0).ToString();
								sEstacion = drBloqueo.GetValue(1).ToString();
								dtFechaBloqueo = drBloqueo.GetDateTime(2);
							}
							drBloqueo.Close();
						
							MessageBox.Show(string.Format("La Nota De Credito Del Cliente esta siendo Editada : \n\n Usuario : '{0}' \n Estacion : '{1}' \n Desde : '{2}' \n\n La Factura Se Cancelara hasta que el usuario termine de Editar la Nota de Credito", sUsuario, sEstacion, dtFechaBloqueo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.txtNumNotaCredito.Focus();
							return;
						}

						this.DialogResult = DialogResult.OK;
					}
					else
					{
						dr.Close();
						MessageBox.Show("El documento no existe o no esta autorizado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtNumNotaCredito.Focus();
						return;
					}
				}
				else
				{
					MessageBox.Show("Ingrese el número de documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumNotaCredito.Focus();
					return;
				}
			}
			else 
				this.DialogResult = DialogResult.OK;
		}

		private void frmFormasDePago_Load(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FormasDePagoVentas {0}, {1}, {2}, {3}, {4}", 
				IdTipoFactura, MenuLatinium.iNivel, bFDPCredito, bFDPContraEntrega, iBodega));

			this.ultraGrid1.Focus();
			this.ultraGrid1.ActiveRow = this.ultraGrid1.Rows[0];
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmFormasDePago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
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
			IdFormaPago = (int) this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;

			if ((int) this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 4)
			{				
				this.txtNumNotaCredito.Enabled = true;
				this.txtNumNotaCredito.Focus();
			}	
			else
			{
				this.txtNumNotaCredito.Enabled = false;
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			IdFormaPago = (int) this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;

			if ((int) this.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 4)
			{				
				this.txtNumNotaCredito.Enabled = true;
				this.txtNumNotaCredito.Focus();
			}	
			else
			{
				this.txtNumNotaCredito.Enabled = false;
			}
		}

		private void txtNumNotaCredito_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumNotaCredito_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNumNotaCredito.Text.Length > 0)
			{
				char Pad = '0';
				this.txtNumNotaCredito.Text = this.txtNumNotaCredito.Text.ToString().PadLeft(9, Pad);
			}
		}

		private void txtNumNotaCredito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnAceptar.Focus();
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}
	}
}

