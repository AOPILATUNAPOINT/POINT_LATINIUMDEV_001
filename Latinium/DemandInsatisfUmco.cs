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
	/// Descripción breve de DemandInsatisfUmco.
	/// </summary>
	public class DemandInsatisfUmco : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.Misc.UltraButton btLeer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.Misc.UltraButton btnExportar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;

		
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DemandInsatisfUmco()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.btLeer = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btnExportar = new Infragistics.Win.Misc.UltraButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			appearance6.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(0, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(712, 376);
			this.ultraGrid1.TabIndex = 16;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// chkCliente
			// 
			this.chkCliente.EditValue = true;
			this.chkCliente.Location = new System.Drawing.Point(40, 24);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkCliente.Properties.Caption = "Todo Cliente";
			this.chkCliente.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.SystemColors.Control, System.Drawing.SystemColors.ControlText);
			this.chkCliente.Size = new System.Drawing.Size(96, 21);
			this.chkCliente.TabIndex = 54;
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataMember = "Cliente";
			this.cmbCliente.DataSource = this.cdsCliente;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.MaxWidth = 100;
			ultraGridColumn1.MinWidth = 100;
			ultraGridColumn1.Width = 100;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.MaxWidth = 100;
			ultraGridColumn3.MinWidth = 100;
			ultraGridColumn3.Width = 300;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
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
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Location = new System.Drawing.Point(152, 24);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(192, 21);
			this.cmbCliente.TabIndex = 0;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// btExcel
			// 
			this.btExcel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btExcel.Location = new System.Drawing.Point(504, 24);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(96, 24);
			this.btExcel.TabIndex = 57;
			this.btExcel.Text = "&Cotizacion";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btLeer
			// 
			this.btLeer.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btLeer.Location = new System.Drawing.Point(376, 24);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(104, 24);
			this.btLeer.TabIndex = 56;
			this.btLeer.Text = "&Generar Reporte";
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// btnExportar
			// 
			this.btnExportar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.btnExportar.Location = new System.Drawing.Point(616, 24);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(96, 24);
			this.btnExportar.TabIndex = 58;
			this.btnExportar.Text = "&Excel";
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// DemandInsatisfUmco
			// 
			this.AcceptButton = this.btLeer;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(728, 453);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btLeer);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "DemandInsatisfUmco";
			this.Text = "DemandInsatisfUmco";
			this.Load += new System.EventHandler(this.DemandInsatisfUmco_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSetU= new DataSet();
		DataView dvG = new DataView();
		private void btLeer_Click(object sender, System.EventArgs e)
		{
			llenaGrid();
		}

		private void DemandInsatisfUmco_Load(object sender, System.EventArgs e)
		{
		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString=			 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbCliente.Enabled = !this.chkCliente.Checked;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			int Id = 0;			
			try
			{	
				Id = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			}
			catch(Exception ex)
			{
				MessageBox.Show("No existe datos para Generar: " + ex.Message,"Información",MessageBoxButtons.OK,MessageBoxIcon.Information);	
				return;
			}			
			generaCotizacion();				
		}
		private void generaCotizacion()
		{
			if (DialogResult.No == MessageBox.Show("Desea Generar una Cotizacion", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			if(verificaFilas()<=1)
			{MessageBox.Show("No Selecciono fila ","Información",MessageBoxButtons.OK,MessageBoxIcon.Information);	
				return;
			}						
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = DateTime.Today;
			string stFactura = "Exec FacturaCrear '', '', '-1', '" + dtFecha.ToString("yyyyMMdd")
				+ "', 0, 0, 0, 0, 0, 3";
			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stFactura, true);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Verifica"].Value==true)
				{
					int idart = (int) dr.Cells["idArticulo"].Value;					
  				int idDC = (int) dr.Cells["idDetCompra"].Value;
					int idCli = (int) dr.Cells["idCliente"].Value;
					double Pendiente = (double) dr.Cells["Pendiente"].Value;
					string stDetCompra = "Exec detCompraUmco "+IdCompra+", "+idart+", "+idDC+", "+idCli+", "+Pendiente;
					Funcion.EjecutaSQL(cdsCliente, stDetCompra, true);													
				}
			}
			llenaGrid();
		Cursor = Cursors.Default;
		
			if (DialogResult.No == MessageBox.Show("Cotizacion Generada \n Desea Revisar Cotizacion", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

//			Compra miCompra1 = new Compra(3);
//			miCompra1.MdiParent = this.MdiParent;
//			miCompra1.Show();

		}
		private int verificaFilas()
		{
			int cont=1;	
			bool verifica;		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{		//cont=1;			
				try 
				{
					verifica = (bool)dr.Cells["Verifica"].Value;
					if(verifica == true)
					{		cont = cont+cont;}
				}
				catch(Exception ex)
				{	Console.Write(ex);	}				
			}
			return cont;
		}

		private void llenaGrid()
		{
			this.errorProvider.SetError(this.cmbCliente,"");
			int chek = 1;
			Object idcliente = cmbCliente.Value;
			if(chkCliente.Checked==true)
			{
				chek = 0;
				idcliente=0;
			}
			else
			{
				if (this.cmbCliente.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbCliente, "Seleccione un Cliente");
					return;
				}
			}

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			
			string stProcedimiento = "exec ReporteDemandaUMCO "+idcliente+" ,"+ chek;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetU.Reset();
			miDataSetU.Clear();
			try
			{
				myData.Fill(miDataSetU);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al cargar Archivo:___ " );
				return;
			}
			dvG = miDataSetU.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;		
			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idDetCompra"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idCliente"].Hidden = true;
			
			#endregion Visibilidad
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "DemandaI.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "DemandaI.xls");

		}

	
	}
}
