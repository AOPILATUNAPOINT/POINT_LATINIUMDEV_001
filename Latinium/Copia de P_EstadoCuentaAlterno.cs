using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de P_EstadoCuentaAlterno.
	/// </summary>
	public class P_EstadoCuentaAlterno : System.Windows.Forms.Form
	{
		#region Variables Sistema
		private Infragistics.Win.Misc.UltraButton btImprimir;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.C1Excel.C1XLBook c1XLBook1;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.Misc.UltraButton btSubir;
		private C1.Data.C1DataSet cdsConciliacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.Misc.UltraButton btLeer;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.ComponentModel.IContainer components;
		#endregion Variables Sistema
		#region Construtor
		public P_EstadoCuentaAlterno()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		string idcuenta="";
		public P_EstadoCuentaAlterno(string VidCuenta)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idcuenta = VidCuenta;

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
#endregion Construtor
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
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btSubir = new Infragistics.Win.Misc.UltraButton();
			this.btLeer = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.c1XLBook1 = new C1.C1Excel.C1XLBook();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cdsConciliacion = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConciliacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImprimir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImprimir.Location = new System.Drawing.Point(192, 296);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(64, 24);
			this.btImprimir.TabIndex = 23;
			this.btImprimir.Text = "&Imprimir";
			this.toolTip1.SetToolTip(this.btImprimir, "Reporte no conciliados");
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAsiento.Location = new System.Drawing.Point(104, 296);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(64, 24);
			this.btAsiento.TabIndex = 22;
			this.btAsiento.Text = "&Asiento";
			this.toolTip1.SetToolTip(this.btAsiento, "Contabilizar los marcados con visto");
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btSubir
			// 
			this.btSubir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSubir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSubir.Location = new System.Drawing.Point(16, 296);
			this.btSubir.Name = "btSubir";
			this.btSubir.Size = new System.Drawing.Size(64, 24);
			this.btSubir.TabIndex = 24;
			this.btSubir.Text = "&Subir";
			this.toolTip1.SetToolTip(this.btSubir, "Carga Los archivos en Pantalla a Preconciliacion");
			this.btSubir.Visible = false;
			this.btSubir.Click += new System.EventHandler(this.btSubir_Click);
			// 
			// btLeer
			// 
			this.btLeer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btLeer.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btLeer.Location = new System.Drawing.Point(16, 296);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(64, 24);
			this.btLeer.TabIndex = 47;
			this.btLeer.Text = "&Leer";
			this.toolTip1.SetToolTip(this.btLeer, "Busca el archivo y carga en la pantalla");
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(104, 296);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 24);
			this.btCancelar.TabIndex = 48;
			this.btCancelar.Text = "&Cancelar";
			this.toolTip1.SetToolTip(this.btCancelar, "Cancela la subida de archivos");
			this.btCancelar.Visible = false;
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrabar.Location = new System.Drawing.Point(280, 296);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 24);
			this.btGrabar.TabIndex = 49;
			this.btGrabar.Text = "&Borrar";
			this.toolTip1.SetToolTip(this.btGrabar, "Borrar Seleccionados");
			this.btGrabar.Visible = false;
			// 
			// cdsConciliacion
			// 
			this.cdsConciliacion.BindingContextCtrl = this;
			this.cdsConciliacion.DataLibrary = "LibContabilidad";
			this.cdsConciliacion.DataLibraryUrl = "";
			this.cdsConciliacion.DataSetDef = "dsTmpConcilia";
			this.cdsConciliacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsConciliacion.Name = "cdsConciliacion";
			this.cdsConciliacion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsConciliacion.SchemaDef = null;
			this.cdsConciliacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsConciliacion_BeforeFill);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance4.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(12, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(792, 280);
			this.ultraGrid1.TabIndex = 26;
			// 
			// chkTodos
			// 
			this.chkTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTodos.Appearance = appearance7;
			this.chkTodos.Location = new System.Drawing.Point(624, 296);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(184, 24);
			this.chkTodos.TabIndex = 45;
			this.chkTodos.Text = "Confirmados / No Confirmados";
			this.chkTodos.Visible = false;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			appearance8.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance8;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(792, 280);
			this.ultraGrid2.TabIndex = 46;
			// 
			// cmbCuenta
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance14;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.cmbCuenta.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.cmbCuenta.DisplayMember = "";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(432, 296);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(192, 21);
			this.cmbCuenta.TabIndex = 50;
			this.cmbCuenta.ValueMember = "";
			this.cmbCuenta.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbCuenta_MouseDown);
			this.cmbCuenta.Click += new System.EventHandler(this.cmbCuenta_Click);
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(368, 304);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 51;
			this.label3.Text = "Cuenta Nº";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// P_EstadoCuentaAlterno
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(816, 333);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btLeer);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btSubir);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btAsiento);
			this.Name = "P_EstadoCuentaAlterno";
			this.Text = "Pre Conciliacion";
			this.Load += new System.EventHandler(this.P_EstadoCuentaAlterno_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsConciliacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		#region Inicializacion
		private void P_EstadoCuentaAlterno_Load(object sender, System.EventArgs e)
		{
			ultraGrid1.Visible=true;
			ultraGrid2.Visible= false;
			LlenaPreconciliacion(0);
			llenaComboCuenta();
		}
		private void cdsConciliacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsConciliacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		#endregion Inicializacion
		#region Botones de pantalla
		private void btLeer_Click(object sender, System.EventArgs e)
		{						
			int i = AbrirArchivo();
			if (i==1)return;
			LeerDatosExcel();				
		}
		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkTodos.Checked==true)
			{
				btGrabar.Visible=true;
				btAsiento.Enabled=false;				
				LlenaPreconciliacion(1);
			}
			else
			{
				btGrabar.Visible=false;
				btAsiento.Enabled=true;	
				LlenaPreconciliacion(0);
			}
		
		}
		private void btSubir_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Subir PreConciliacion de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
			{
				visibles(2);
				return;
			}
			Subir();
			LlenaPreconciliacion(0);			
			visibles(2);
		}	
		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			visibles(2);
		}
		private void btAsiento_Click(object sender, System.EventArgs e)
		{
		 CreaAsiento();
		}			
		private void btImprimir_Click(object sender, System.EventArgs e)
		{			
			this.Cursor = Cursors.WaitCursor;		
			string stTitulo = "REPORTE NO CONFIRMADOS";
			string stFiltro = " {p_Preconciliacion.idCuenta} = "+idcuenta;								
			Reporte miReporte = new Reporte("P_PreConciliacion.rpt", stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(stTitulo);
			miReporte.Show();
			this.Cursor = Cursors.Default;
		}
			
		private void cmbCuenta_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right)return;			 
				string stPrimero = "Descripcion";
				if (cmbCuenta.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
					stPrimero = "CodRapido";			
				cmbCuenta.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
				cmbCuenta.DisplayMember = stPrimero;
			
		}
		private void cmbCuenta_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbCuenta, "");
		}
		#endregion Botones de pantalla
		#region Funciones Personalizadas		
		//Declaracion de Variables Globales
		string stArchivo = "";
		DataSet miDataSet = new DataSet();
		private int  AbrirArchivo()
		{
			int retorna = 0;
			this.openFileDialog.Filter = "xls (*.xls)|*.xls";						
			this.openFileDialog.InitialDirectory = Application.ExecutablePath;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();				
			}
			else
			{
				MessageBox.Show("Cancelado por el usuario");
				visibles(2);
				this.Cursor = Cursors.Default;				
				retorna = 1;
			}
		return retorna;
		}
		
		private void LeerDatosExcel()
		{
			//boton de facturas para subir del excel a la grilla
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + stArchivo + ";" +
				"Extended Properties='Excel 8.0;'";		
			string stHoja = "select * from [Hoja1$]";
			OleDbDataAdapter myData = new OleDbDataAdapter(stHoja, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener cerrado el archivo " + stArchivo + " en Excel",
					ex.Message);
				visibles(2);		
				return;
			}
			ultraGrid2.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid2.DataBind();
			visibles(1);		
		}
		
		private void Subir()
		{
			Cursor = Cursors.WaitCursor;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{		
				try
				{													
					float debe = float.Parse (dr.Cells[0].Value.ToString());
					float haber	= float.Parse (dr.Cells[1].Value.ToString());
					DateTime dtDesde = DateTime.Parse(dr.Cells[2].Value.ToString());					
					string Referencia	= dr.Cells[3].Value.ToString();
					string notas	= dr.Cells[4].Value.ToString();																		
					string stExec = string.Format("exec P_PreConciliacionP {0},{1},'{2}','{3}','{4}',{5}",
						debe,haber,dtDesde.ToString("yyyyMMdd"),Referencia,notas,idcuenta);
																					
						Funcion.EjecutaSQL(cdsConciliacion,stExec);  
					}
					catch(Exception ex)
					{
						MessageBox.Show("Error al importar Pre Conciliacion...." + ex.Message);
					}

			}
			Cursor = Cursors.Default;
			MessageBox.Show("Proceso Terminado...");
		}
		
		private void visibles(int control)
		{
			chkTodos.Checked = false;
			if(control == 1)
			{
				ultraGrid1.Visible= false;
				ultraGrid2.Visible= true;						
				btAsiento.Visible=false;
				btImprimir.Visible=false;
				btLeer.Visible=false;
				btSubir.Visible=true;
				btCancelar.Visible=true;
//				btGrabar.Visible=false;
			//	chkTodos.Visible=false;
				this.Text = "Datos para subir a Pre Conciliacion";
			}
			else
			{
				ultraGrid1.Visible= true;
				ultraGrid2.Visible= false;			
				btAsiento.Visible=true;
				btImprimir.Visible=true;
				btLeer.Visible=true;
				btSubir.Visible=false;
				btCancelar.Visible=false;
				//btGrabar.Visible=true;
				//chkTodos.Visible=true;
				this.Text = "Pantalla de Pre Conciliacion";
			}
		}

		private void LlenaPreconciliacion(int confirmado)
		{
			string stSelect ="Select * from P_Preconciliacion where confirmar = "+confirmado+" and idCuenta = "+idcuenta;
			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsConciliacion,stSelect);			
			ultraGrid1.Rows.Band.Columns["idPreConciliacion"].Hidden = true;		
			ultraGrid1.Rows.Band.Columns["IdCuenta"].Hidden = true;	
		}
		
		private void llenaComboCuenta()
		{			 
			//combo Cuenta
			string stSelect ="select idCuenta,CodRapido,Descripcion from cuenta where grupo = 0 order by codigo";
			cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsConciliacion,stSelect);			
			cmbCuenta.DisplayMember = "CodRapido";
			cmbCuenta.ValueMember = "idCuenta";
			cmbCuenta.Rows.Band.Columns["idCuenta"].Hidden = true;
			//llena el valor predefinido en la tabla cuenta
			stSelect = "Select isnull(idCuentaPreconciliacion,0) from cuenta where idcuenta = "+idcuenta;
			int idCuentaHaber = Funcion.iEscalarSQL(cdsConciliacion,stSelect);
			cmbCuenta.Value = idCuentaHaber;
		}
		int idAsientoN =0;
		private void CreaAsiento()
		{			
			Cursor = Cursors.WaitCursor;
			errorProvider.SetError(cmbCuenta, "");
			errorProvider.SetError(btAsiento, "");

			if (cmbCuenta.Value == null || cmbCuenta.Value == System.DBNull.Value || (int) cmbCuenta.Value <= 0)
			{
				errorProvider.SetError(cmbCuenta, "Debe Seleccionar una Cuenta para generar el Asiento...");
				Cursor = Cursors.Default;
				return;
			}
			//control para crear asiento si esta seleccionada alguna fila con visto en confirmar
			int i=0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{	
				if (bool.Parse(dr.Cells[1].Value.ToString())== true)
				{
					i++;
				}
			}
			if (i==0)
			{
				errorProvider.SetError(btAsiento, "No selecciono ninguna fila para generar Asiento...");
				Cursor = Cursors.Default;
				return;
			}
			else
			{
				try
				{
					string stExec = string.Format("exec P_PreConciliacionAsiento ");
					idAsientoN = Funcion.iEscalarSQL(cdsConciliacion,stExec);
					errorProvider.SetError(btAsiento, "");
				}
				catch(Exception ex)
				{
					MessageBox.Show("No se pudo crear el asiento..."+ex);
					Cursor = Cursors.Default;
					return;
				}
			}

		// proceso para insertar lineas en el detalle del asiento
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{		
				try
				{	
					if (bool.Parse(dr.Cells[1].Value.ToString())== true)
					{
						int idPreC = int.Parse(dr.Cells[0].Value.ToString());
						float debe = float.Parse (dr.Cells[2].Value.ToString());
						float haber	= float.Parse (dr.Cells[3].Value.ToString());
						DateTime dtDesde = DateTime.Parse(dr.Cells[4].Value.ToString());					
						string Referencia	= dr.Cells[5].Value.ToString();
						string notas	= dr.Cells[6].Value.ToString();																		
						string stExec = string.Format("exec P_PreconciliacionP {0},{1},'{2}','{3}','{4}',{5},{6},{7},{8}",
							debe,haber,dtDesde.ToString("yyyyMMdd"),Referencia,notas,idcuenta,idPreC,idAsientoN,cmbCuenta.Value);																					
						Funcion.EjecutaSQL(cdsConciliacion,stExec);
					}
				}
				catch(Exception ex)
				{
					MessageBox.Show("Error al Crear Asiento...." + ex.Message);
					Cursor = Cursors.Default;
					return;
				}
			}

			//inserta la ultima linea
			string stExec1 = string.Format("exec P_PreconciliacionP {0},{1},'{2}','{3}','{4}',{5},{6},{7},{8},{9}",
			0,0,DateTime.Today.ToString("yyyyMMdd"),"","",0,0,idAsientoN,cmbCuenta.Value,1);																					
			Funcion.EjecutaSQL(cdsConciliacion,stExec1);
			LlenaPreconciliacion(0);
			Cursor = Cursors.Default;

		}
		#endregion Funciones Personalizadas

	
	

		


	

		
	}
}
