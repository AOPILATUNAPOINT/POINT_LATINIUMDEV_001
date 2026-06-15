using System.Data;
using System.Data.SqlClient;using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de IngresoContratosBuscar.
	/// </summary>
	public class IngresoContratosBuscar : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaH;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaD;
		private Infragistics.Win.Misc.UltraButton btAbrir;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private DevExpress.XtraEditors.CheckEdit chkCliente;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCliente;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		private DevExpress.XtraEditors.CheckEdit chkLocal;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public IngresoContratosBuscar()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFondo");
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbFechaH = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaD = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btAbrir = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.chkCliente = new DevExpress.XtraEditors.CheckEdit();
			this.cdvCliente = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkLocal = new DevExpress.XtraEditors.CheckEdit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLocal.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(304, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 24);
			this.label3.TabIndex = 51;
			this.label3.Text = "Fecha Hasta";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(24, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 24);
			this.label2.TabIndex = 50;
			this.label2.Text = "Fecha Desde";
			// 
			// cmbFechaH
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaH.Appearance = appearance1;
			this.cmbFechaH.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaH.DateButtons.Add(dateButton1);
			this.cmbFechaH.Format = "dd/MMM/yyyy";
			this.cmbFechaH.Location = new System.Drawing.Point(392, 8);
			this.cmbFechaH.Name = "cmbFechaH";
			this.cmbFechaH.NonAutoSizeHeight = 23;
			this.cmbFechaH.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaH.SpinButtonsVisible = true;
			this.cmbFechaH.TabIndex = 49;
			this.cmbFechaH.Value = new System.DateTime(2009, 10, 26, 0, 0, 0, 0);
			// 
			// cmbFechaD
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaD.Appearance = appearance2;
			this.cmbFechaD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaD.DateButtons.Add(dateButton2);
			this.cmbFechaD.Format = "dd/MMM/yyyy";
			this.cmbFechaD.Location = new System.Drawing.Point(120, 8);
			this.cmbFechaD.Name = "cmbFechaD";
			this.cmbFechaD.NonAutoSizeHeight = 23;
			this.cmbFechaD.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaD.SpinButtonsVisible = true;
			this.cmbFechaD.TabIndex = 48;
			this.cmbFechaD.Value = new System.DateTime(2009, 10, 26, 0, 0, 0, 0);
			// 
			// btAbrir
			// 
			this.btAbrir.Enabled = false;
			this.btAbrir.Location = new System.Drawing.Point(424, 72);
			this.btAbrir.Name = "btAbrir";
			this.btAbrir.Size = new System.Drawing.Size(80, 24);
			this.btAbrir.TabIndex = 47;
			this.btAbrir.Text = "Ir...";
			this.btAbrir.Click += new System.EventHandler(this.btAbrir_Click);
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
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
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
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(0, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(504, 280);
			this.ultraGrid1.TabIndex = 46;
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(328, 72);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 24);
			this.btBuscar.TabIndex = 45;
			this.btBuscar.Text = "Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
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
			ultraGridColumn10.Header.VisiblePosition = 9;
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
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Location = new System.Drawing.Point(120, 40);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(176, 21);
			this.cmbCliente.TabIndex = 55;
			this.cmbCliente.ValueMember = "idCliente";
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
			// chkCliente
			// 
			this.chkCliente.EditValue = true;
			this.chkCliente.Location = new System.Drawing.Point(8, 40);
			this.chkCliente.Name = "chkCliente";
			// 
			// chkCliente.Properties
			// 
			this.chkCliente.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkCliente.Properties.Caption = "Todo Cliente";
			this.chkCliente.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.SystemColors.Control, System.Drawing.SystemColors.ControlText);
			this.chkCliente.Size = new System.Drawing.Size(96, 21);
			this.chkCliente.TabIndex = 56;
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cmbLocal
			// 
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.cmbLocal.Enabled = false;
			this.cmbLocal.Location = new System.Drawing.Point(120, 72);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(176, 21);
			this.cmbLocal.TabIndex = 57;
			this.cmbLocal.ValueMember = "";
			// 
			// chkLocal
			// 
			this.chkLocal.EditValue = true;
			this.chkLocal.Location = new System.Drawing.Point(8, 72);
			this.chkLocal.Name = "chkLocal";
			// 
			// chkLocal.Properties
			// 
			this.chkLocal.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.chkLocal.Properties.Caption = "Todo Local";
			this.chkLocal.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.SystemColors.Control, System.Drawing.SystemColors.ControlText);
			this.chkLocal.Size = new System.Drawing.Size(96, 21);
			this.chkLocal.TabIndex = 58;
			this.chkLocal.CheckedChanged += new System.EventHandler(this.chkLocal_CheckedChanged);
			// 
			// IngresoContratosBuscar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(512, 357);
			this.Controls.Add(this.cmbLocal);
			this.Controls.Add(this.chkLocal);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbFechaH);
			this.Controls.Add(this.cmbFechaD);
			this.Controls.Add(this.btAbrir);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btBuscar);
			this.Name = "IngresoContratosBuscar";
			this.Text = "IngresoContratosBuscar";
			this.Load += new System.EventHandler(this.IngresoContratosBuscar_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkLocal.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbFechaD.Value;
			DateTime dtHasta = (DateTime) cmbFechaH.Value;									

			this.errorProvider.SetError(this.cmbCliente,"");
			this.errorProvider.SetError(this.cmbLocal,"");
			Object idcliente = cmbCliente.Value;
			if(chkCliente.Checked==true)
			{
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
			int idLocales=0;
			if(chkLocal.Checked!=true)
			{
				if (this.cmbLocal.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbLocal, "Seleccione un Local");
					return;
				}
				idLocales = int.Parse(cmbLocal.Value.ToString());
			}

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec ContratosBuscar '"+dtDesde.ToString("yyyyMMdd")+"', '"+dtHasta.ToString("yyyyMMdd")+" ', '"+idcliente+" ', '"+idLocales+"'";;
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
public int idBusca=0;
		private void btAbrir_Click(object sender, System.EventArgs e)
		{
			if(ultraGrid1.Rows.Count < 1 )return;
			idBusca = (int) ultraGrid1.ActiveRow.Cells["idContratos"].Value;
			if(idBusca<0)return;
			this.DialogResult = DialogResult.OK;
			this.Close();

		}

		private void IngresoContratosBuscar_Load(object sender, System.EventArgs e)
		{
			string stSelect ="select idLocales,Nombre from contratoslocales order by nombre";
			cmbLocal.DataSource = Funcion.dvProcedimiento(cdsCliente,stSelect);
			
			cmbLocal.Rows.Band.Columns["idLocales"].Hidden = true;
			cmbLocal.DisplayMember = "Nombre";
			cmbLocal.ValueMember = "idLocales";

			DateTime dtDesde = DateTime.Today.AddMonths(-1).AddDays(-DateTime.Today.Day+1);
			
			cmbFechaD.Value = dtDesde;
			cmbFechaH.Value = dtDesde.AddMonths(1).AddDays(-1);
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbCliente.Enabled = !this.chkCliente.Checked;

		}

		private void chkLocal_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbLocal.Enabled = !this.chkLocal.Checked;

		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
	}
}
