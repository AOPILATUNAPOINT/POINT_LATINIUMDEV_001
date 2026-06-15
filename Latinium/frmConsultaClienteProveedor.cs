using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConsultaClienteProveedor.
	/// </summary>
	public class frmConsultaClienteProveedor : System.Windows.Forms.Form
	{
		private System.Data.DataView dataView1;
		public Infragistics.Win.UltraWinGrid.UltraGrid udtG;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnBusca;
		private System.Windows.Forms.Button btnAcepta;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtid;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConsultaClienteProveedor()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		bool bClienteProveedor = false;

		public frmConsultaClienteProveedor(bool BClienteProveedor)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			bClienteProveedor = BClienteProveedor;

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("clg.Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agente_de_retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sujeto_de_retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("clg.Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Agente_de_retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sujeto_de_retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.dataView1 = new System.Data.DataView();
			this.udtG = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.btnBusca = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAcepta = new System.Windows.Forms.Button();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtid = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.dataView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udtG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtid)).BeginInit();
			this.SuspendLayout();
			// 
			// udtG
			// 
			this.udtG.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udtG.Cursor = System.Windows.Forms.Cursors.Default;
			this.udtG.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.udtG.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "Cliente";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Grupo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn7.Header.Caption = "Agente de retencion";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.Caption = "Sujeto de retención";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
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
			this.udtG.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.udtG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udtG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udtG.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.udtG.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udtG.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.udtG.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udtG.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udtG.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.udtG.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.udtG.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.udtG.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.udtG.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.udtG.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.udtG.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.udtG.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.udtG.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udtG.Location = new System.Drawing.Point(0, 48);
			this.udtG.Name = "udtG";
			this.udtG.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.udtG.Size = new System.Drawing.Size(856, 192);
			this.udtG.TabIndex = 38;
			this.udtG.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.udtG_KeyPress);
			this.udtG.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udtG_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
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
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(24, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 290;
			this.label1.Text = "Nombre/Ruc";
			// 
			// btnBusca
			// 
			this.btnBusca.Location = new System.Drawing.Point(272, 13);
			this.btnBusca.Name = "btnBusca";
			this.btnBusca.TabIndex = 291;
			this.btnBusca.Text = "&Buscar";
			this.btnBusca.Click += new System.EventHandler(this.btnBusca_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Location = new System.Drawing.Point(120, 248);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 292;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.button2_Click);
			// 
			// btnAcepta
			// 
			this.btnAcepta.Location = new System.Drawing.Point(32, 248);
			this.btnAcepta.Name = "btnAcepta";
			this.btnAcepta.TabIndex = 293;
			this.btnAcepta.Text = "&Aceptar";
			this.btnAcepta.Click += new System.EventHandler(this.btnAcepta_Click);
			// 
			// txtNombre
			// 
			appearance6.BackColorDisabled = System.Drawing.Color.Silver;
			appearance6.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance6.BorderColor = System.Drawing.Color.Black;
			appearance6.BorderColor3DBase = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance6;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(112, 14);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(120, 21);
			this.txtNombre.TabIndex = 294;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraTextEditor1_KeyDown);
			// 
			// txtid
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtid.Appearance = appearance7;
			this.txtid.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtid.Enabled = false;
			this.txtid.FormatString = "0";
			this.txtid.Location = new System.Drawing.Point(40, 424);
			this.txtid.MinValue = 0;
			this.txtid.Name = "txtid";
			this.txtid.PromptChar = ' ';
			this.txtid.Size = new System.Drawing.Size(8, 21);
			this.txtid.TabIndex = 295;
			this.txtid.Visible = false;
			// 
			// frmConsultaClienteProveedor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(864, 278);
			this.Controls.Add(this.txtid);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.btnAcepta);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnBusca);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.udtG);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmConsultaClienteProveedor";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "frmConsultaClienteProveedor";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConsultaClienteProveedor_KeyDown);
			this.Load += new System.EventHandler(this.frmConsultaClienteProveedor_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udtG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtid)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
				
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}

		private void udtG_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
     private void txtRucId_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validarN.Solonumeros(e);
		}

		private void Aceptar()
		{
			if (this.udtG.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros encontrados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtNombre.Focus();			
			}
			else if (this.udtG.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (this.udtG.ActiveRow.Cells["idCliente"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else this.DialogResult = DialogResult.OK;
		}
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void frmConsultaClienteProveedor_Load(object sender, System.EventArgs e)
		{
			
			if(!bClienteProveedor)
			{
				this.txtid.Value= 0;
				this.Text = "CONSULTAR CLIENTE";
				this.udtG.DisplayLayout.Bands[0].Columns["Agente_de_retencion"].Hidden = false;
				this.udtG.DisplayLayout.Bands[0].Columns["Sujeto_de_retencion"].Hidden = true;
			}
		
			if(bClienteProveedor)
			{
				this.txtid.Value= 1;
				this.Text = "CONSULTAR PROVEEDOR";
				this.udtG.DisplayLayout.Bands[0].Columns["Agente_de_retencion"].Hidden = true;
				this.udtG.DisplayLayout.Bands[0].Columns["Sujeto_de_retencion"].Hidden = false;

			}
			

		}

		private void udtG_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) Aceptar();
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void simpleButton1_Click_1(object sender, System.EventArgs e)
		{
			
		}

		private void btnBusca_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaGeneralClienteProveedor {0}, '{1}'", 
				this.txtid.Value,this.txtNombre.Text.ToString());
			this.udtG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			if (this.udtG.Rows.Count > 0) this.udtG.Focus(); else this.txtNombre.Focus();
		}

		private void btnAcepta_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmConsultaClienteProveedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Escape)
				this.DialogResult = DialogResult.No;
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Enter)
				this.btnBusca_Click(sender,e);
		}

		private void ultraTextEditor1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Enter)
				this.btnBusca_Click(sender,e);
		}
	}

	class validarN 
	{
		public static void Solonumeros (KeyPressEventArgs Nu)
		{
			if(char.IsDigit(Nu.KeyChar))
			{
			
				Nu.Handled = false;
			
			}
			else if (char.IsControl(Nu.KeyChar))
			{
			
				Nu.Handled = false;			} 
			else
			{
				Nu.Handled= true;
			}
		}
	
	
	}
}
