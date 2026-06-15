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
	/// Descripción breve de IngersoContratosFacturas.
	/// </summary>
	public class PresupuestoPedidos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private System.Windows.Forms.ImageList imageList;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCompra;
		private System.ComponentModel.IContainer components;

		public PresupuestoPedidos()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PresupuestoPedidos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(496, 288);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(112, 24);
			this.btAceptar.TabIndex = 11;
			this.btAceptar.Text = "Generar &Factura";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btFiltro
			// 
			appearance1.Image = 5;
			this.btFiltro.Appearance = appearance1;
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.ImageList = this.imageList;
			this.btFiltro.Location = new System.Drawing.Point(336, 16);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(24, 24);
			this.btFiltro.TabIndex = 12;
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
			this.imageList.ImageSize = new System.Drawing.Size(18, 18);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(624, 288);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 24);
			this.btCancelar.TabIndex = 17;
			this.btCancelar.Text = "Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbCliente
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCliente.Appearance = appearance2;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 325;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(72, 16);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(248, 21);
			this.cmbCliente.TabIndex = 18;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(64, 16);
			this.label1.TabIndex = 19;
			this.label1.Text = "Proveedor";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			appearance8.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance8;
			this.ultraGrid1.Location = new System.Drawing.Point(0, 48);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(712, 232);
			this.ultraGrid1.TabIndex = 20;
			// 
			// PresupuestoPedidos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(712, 317);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "PresupuestoPedidos";
			this.Text = "Generar Ingresos de Bodega";
			this.Load += new System.EventHandler(this.PresupuestoPedidos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void PresupuestoPedidos_Load(object sender, System.EventArgs e)
		{
			string stSelect ="Select idcliente,Nombre,Codigo from Cliente where proveedor = 1 or ambos = 1 order by nombre";
			cmbCliente.DataSource = Funcion.dvProcedimiento(cdsCompra,stSelect);			
			cmbCliente.Rows.Band.Columns["idcliente"].Hidden = true;
			cmbCliente.DisplayMember = "Nombre";
			cmbCliente.ValueMember = "idcliente";
		
		}
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
	
		private void btFiltro_Click(object sender, System.EventArgs e)
		{	
			errorProvider.SetError(cmbCliente, "");			
			if (cmbCliente.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbCliente, "Ingrese Obra");
				return;
			}

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec P_PresupuestoFiltraOrdenCompra "+cmbCliente.Value;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;
			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idDetCompra"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;			
			ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;		
			ultraGrid1.DisplayLayout.Bands[0].Columns["idReporte"].Hidden = true;		
			#endregion Visibilidad			
		}
		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
		}
//Boton genera facturas
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(btAceptar,"");
			if(ultraGrid1.Rows.Count < 1)
			{
				errorProvider.SetError(this.btAceptar,"No Hay datos para Generar....");
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Generar Ingresos de Bodega", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;	
																	
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = DateTime.Today;
			int control =0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in ultraGrid1.Rows.All)
			{																				
				//					//detalle de compra
				if (float.Parse(dr2.Cells["CantPasar"].Value.ToString()) > 0 )
				{
					control ++;
				}
			}
			if(control < 1)
			{
				MessageBox.Show("No Seleccionado Valores Para Crear Ingreso");
				Cursor = Cursors.Default;
				return;
			}

			string stFactura = "Exec FacturaCrear '"+cmbCliente.DisplayMember+"', '', '-1', '" + DateTime.Today.ToString("yyyyMMdd") + "', 0, 0, 0, 0, 0, 9";
			int IdCompra = Funcion.iEscalarSQL(cdsCompra, stFactura, true);	
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr3 in ultraGrid1.Rows.All)
			{																				
				//					//detalle de compra
				if (float.Parse(dr3.Cells["CantPasar"].Value.ToString()) > 0 )
				{
					int idCompra1 = (int) dr3.Cells["idCompra"].Value;
					int idArticulo1 = (int) dr3.Cells["idArticulo"].Value;	
					int idDetCompra = (int) dr3.Cells["idDetCompra"].Value;	
					
					if ((double)dr3.Cells["CantPasar"].Value > (double)dr3.Cells["Saldo"].Value)
					{						
						MessageBox.Show("Cantidad "+dr3.Cells["CantPasar"].Value.ToString() +" a pasar no debe ser mayor al Saldo "+ dr3.Cells["Saldo"].Value.ToString(),"Cantidad Mayor al Saldo...",MessageBoxButtons.OK,MessageBoxIcon.Information);					
						Cursor = Cursors.Default;		
						return;
					}

					string stFactTemp = "Exec P_PresupuestoCreaDetalle "+idDetCompra+", "+idCompra1+", "+IdCompra+", "+idArticulo1+" ,"+dr3.Cells["CantPasar"].Value;	
					Funcion.EjecutaSQL(cdsCompra, stFactTemp, true);
				}																									
			}
			btFiltro_Click(this,e);
			Cursor = Cursors.Default;					
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}
	}
}
