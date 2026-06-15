using System.Data.SqlClient;
using System.Data;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ContratosAlarma.
	/// </summary>
	public class ContratosAlarma : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.Misc.UltraButton btAbrir;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private C1.Data.C1DataSet cdsAlarma;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label lblDias;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ContratosAlarma()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			//ContratosAlarma.Size = new System.Drawing.Size(100,100);			
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		public ContratosAlarma(bool alarma)
		{
			InitializeComponent();
			if(alarma == true)
			{
				//panel1.Visible=false;
				ultraGrid1.Visible = true;
				btAbrir.Visible=true;
				label1.Visible=true;
				llenaGrilla();
			}
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
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.cdsAlarma = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.btAbrir = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.panel1 = new System.Windows.Forms.Panel();
			this.lblDias = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			this.panel1.SuspendLayout();
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
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
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
			this.ultraGrid1.Location = new System.Drawing.Point(8, 48);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(688, 0);
			this.ultraGrid1.TabIndex = 18;
			this.ultraGrid1.Visible = false;
			// 
			// chkCliente
			// 
			this.chkCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "ContratosAlarma.Alarma"));
			this.chkCliente.Location = new System.Drawing.Point(16, 8);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.Caption = "Activar Recordatorio";
			this.chkCliente.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.SystemColors.Control, System.Drawing.SystemColors.ControlText);
			this.chkCliente.Size = new System.Drawing.Size(136, 19);
			this.chkCliente.TabIndex = 58;
			// 
			// cdsAlarma
			// 
			this.cdsAlarma.BindingContextCtrl = this;
			this.cdsAlarma.DataLibrary = "LibFacturacion";
			this.cdsAlarma.DataLibraryUrl = "";
			this.cdsAlarma.DataSetDef = "dsContrato";
			this.cdsAlarma.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAlarma.Name = "cdsAlarma";
			this.cdsAlarma.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsAlarma.SchemaDef = null;
			this.cdsAlarma.BeforeFill += new C1.Data.FillEventHandler(this.cdsAlarma_BeforeFill);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(24, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(344, 24);
			this.label1.TabIndex = 59;
			this.label1.Text = "Lista de contratos proximos a Vencer";
			this.label1.Visible = false;
			// 
			// txtAnio
			// 
			this.txtAnio.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsAlarma, "ContratosAlarma.TiempoAlarma"));
			this.txtAnio.Location = new System.Drawing.Point(184, 8);
			this.txtAnio.MaskInput = "nnnnn";
			this.txtAnio.MinValue = 0;
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(64, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 112;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(192, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 16);
			this.label2.TabIndex = 113;
			this.label2.Text = "Dias ";
			// 
			// btAbrir
			// 
			this.btAbrir.Enabled = false;
			this.btAbrir.Location = new System.Drawing.Point(472, 8);
			this.btAbrir.Name = "btAbrir";
			this.btAbrir.Size = new System.Drawing.Size(64, 24);
			this.btAbrir.TabIndex = 114;
			this.btAbrir.Text = "Ir...";
			this.btAbrir.Visible = false;
			this.btAbrir.Click += new System.EventHandler(this.btAbrir_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(288, 8);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(56, 24);
			this.ultraButton1.TabIndex = 115;
			this.ultraButton1.Text = "Grabar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.chkCliente);
			this.panel1.Controls.Add(this.txtAnio);
			this.panel1.Controls.Add(this.label2);
			this.panel1.Controls.Add(this.ultraButton1);
			this.panel1.Location = new System.Drawing.Point(8, 8);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(384, 40);
			this.panel1.TabIndex = 116;
			// 
			// lblDias
			// 
			this.lblDias.Location = new System.Drawing.Point(152, 16);
			this.lblDias.Name = "lblDias";
			this.lblDias.Size = new System.Drawing.Size(40, 16);
			this.lblDias.TabIndex = 117;
			this.lblDias.Text = "Dias";
			// 
			// ContratosAlarma
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(400, 53);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.btAbrir);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblDias);
			this.Controls.Add(this.label1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ContratosAlarma";
			this.Text = "Recordatorio de Contratos";
			this.Load += new System.EventHandler(this.ContratosAlarma_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			this.panel1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
public int idAlarma=0;
		private void ContratosAlarma_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btAbrir_Click(object sender, System.EventArgs e)
		{
			if(ultraGrid1.Rows.Count < 1 )return;
			idAlarma = (int) ultraGrid1.ActiveRow.Cells["idContratos"].Value;
			if(idAlarma<0)return;
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			try
			{
				this.cdsAlarma.Update();
				this.Close();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Grabar");
			}
		}

		private void cdsAlarma_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAlarma.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		C1.Data.C1DataRow drContrato;

		private void llenaGrilla()
		{
			drContrato = cdsAlarma.TableViews["ContratosAlarma"].Rows[0];
	    int tAlm = (int) drContrato["TiempoAlarma"];
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec ContratosPendientes " +tAlm;
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
			ultraGrid1.DisplayLayout.Bands[0].Columns["idContratos"].Hidden = true;
			btAbrir.Enabled=true;
		}
		
	}
}
