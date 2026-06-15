using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de IngresoContratosAmotiza.
	/// </summary>
	public class IngresoContratosAmotiza : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsContratos;
		private System.Windows.Forms.Button btExcel;
		private System.Windows.Forms.Button btGenerar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaP;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaAmortizacion;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoServicio;
		private System.Windows.Forms.Button btMostrar;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button btRecalculo;
		private System.Windows.Forms.Button btBuscar;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContrato;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLocal;
		private System.Windows.Forms.Button btImprimir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public IngresoContratosAmotiza()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btExcel = new System.Windows.Forms.Button();
			this.btGenerar = new System.Windows.Forms.Button();
			this.cdsContratos = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cmbFechaP = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFechaAmortizacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbTipoServicio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btMostrar = new System.Windows.Forms.Button();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.btRecalculo = new System.Windows.Forms.Button();
			this.btBuscar = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.txtContrato = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtLocal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btImprimir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaAmortizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).BeginInit();
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
			this.ultraGrid1.Location = new System.Drawing.Point(648, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(0, 0);
			this.ultraGrid1.TabIndex = 17;
			this.ultraGrid1.Visible = false;
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(552, 80);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(64, 23);
			this.btExcel.TabIndex = 22;
			this.btExcel.Text = "Excel";
			this.btExcel.Visible = false;
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(24, 128);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(64, 23);
			this.btGenerar.TabIndex = 23;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// cdsContratos
			// 
			this.cdsContratos.BindingContextCtrl = this;
			this.cdsContratos.DataLibrary = "LibFacturacion";
			this.cdsContratos.DataLibraryUrl = "";
			this.cdsContratos.DataSetDef = "dsContrato";
			this.cdsContratos.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContratos.Name = "cdsContratos";
			this.cdsContratos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContratos.SchemaDef = null;
			this.cdsContratos.BeforeFill += new C1.Data.FillEventHandler(this.cdsContratos_BeforeFill);
			// 
			// cmbFechaP
			// 
			this.cmbFechaP.AllowMonthSelection = true;
			this.cmbFechaP.AllowWeekSelection = true;
			this.cmbFechaP.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaP.DateButtons.Add(dateButton1);
			this.cmbFechaP.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaP.Format = "dd/MMM/yyyy";
			this.cmbFechaP.Location = new System.Drawing.Point(512, 144);
			this.cmbFechaP.Name = "cmbFechaP";
			this.cmbFechaP.NonAutoSizeHeight = 23;
			this.cmbFechaP.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaP.SpinButtonsVisible = true;
			this.cmbFechaP.TabIndex = 126;
			this.cmbFechaP.Value = new System.DateTime(2010, 5, 28, 0, 0, 0, 0);
			this.cmbFechaP.Visible = false;
			// 
			// ultraLabel14
			// 
			this.ultraLabel14.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel14.Location = new System.Drawing.Point(408, 144);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(96, 16);
			this.ultraLabel14.TabIndex = 127;
			this.ultraLabel14.Text = "Fecha Contratos:";
			this.ultraLabel14.Visible = false;
			// 
			// cmbFechaAmortizacion
			// 
			this.cmbFechaAmortizacion.AllowMonthSelection = true;
			this.cmbFechaAmortizacion.AllowWeekSelection = true;
			this.cmbFechaAmortizacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaAmortizacion.DateButtons.Add(dateButton2);
			this.cmbFechaAmortizacion.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaAmortizacion.Format = "dd/MMM/yyyy";
			this.cmbFechaAmortizacion.Location = new System.Drawing.Point(112, 8);
			this.cmbFechaAmortizacion.Name = "cmbFechaAmortizacion";
			this.cmbFechaAmortizacion.NonAutoSizeHeight = 23;
			this.cmbFechaAmortizacion.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaAmortizacion.SpinButtonsVisible = true;
			this.cmbFechaAmortizacion.TabIndex = 128;
			this.cmbFechaAmortizacion.Value = new System.DateTime(2010, 5, 28, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(8, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(104, 16);
			this.ultraLabel1.TabIndex = 129;
			this.ultraLabel1.Text = "FechaAmortizacion:";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 131;
			this.label1.Text = "Tipo Servicio:";
			// 
			// cmbTipoServicio
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoServicio.Appearance = appearance7;
			this.cmbTipoServicio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoServicio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoServicio.DisplayMember = "";
			this.cmbTipoServicio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoServicio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoServicio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoServicio.Location = new System.Drawing.Point(112, 56);
			this.cmbTipoServicio.Name = "cmbTipoServicio";
			this.cmbTipoServicio.Size = new System.Drawing.Size(112, 21);
			this.cmbTipoServicio.TabIndex = 130;
			this.cmbTipoServicio.ValueMember = "";
			// 
			// btMostrar
			// 
			this.btMostrar.Location = new System.Drawing.Point(560, 112);
			this.btMostrar.Name = "btMostrar";
			this.btMostrar.Size = new System.Drawing.Size(64, 23);
			this.btMostrar.TabIndex = 132;
			this.btMostrar.Text = "Mostrar";
			this.btMostrar.Visible = false;
			this.btMostrar.Click += new System.EventHandler(this.btMostrar_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(328, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 134;
			this.label2.Text = "Cliente";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(328, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(48, 16);
			this.label3.TabIndex = 136;
			this.label3.Text = "Local";
			// 
			// btRecalculo
			// 
			this.btRecalculo.Location = new System.Drawing.Point(104, 128);
			this.btRecalculo.Name = "btRecalculo";
			this.btRecalculo.Size = new System.Drawing.Size(64, 23);
			this.btRecalculo.TabIndex = 137;
			this.btRecalculo.Text = "Recalculo";
			this.btRecalculo.Click += new System.EventHandler(this.btRecalculo_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(248, 32);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(64, 23);
			this.btBuscar.TabIndex = 138;
			this.btBuscar.Text = "Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(328, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(64, 16);
			this.label4.TabIndex = 139;
			this.label4.Text = "Contrato N";
			// 
			// txtContrato
			// 
			this.txtContrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContrato.Location = new System.Drawing.Point(400, 8);
			this.txtContrato.Name = "txtContrato";
			this.txtContrato.Size = new System.Drawing.Size(100, 21);
			this.txtContrato.TabIndex = 140;
			// 
			// txtCliente
			// 
			this.txtCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCliente.Location = new System.Drawing.Point(400, 32);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(216, 21);
			this.txtCliente.TabIndex = 141;
			// 
			// txtLocal
			// 
			this.txtLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLocal.Location = new System.Drawing.Point(400, 56);
			this.txtLocal.Name = "txtLocal";
			this.txtLocal.Size = new System.Drawing.Size(216, 21);
			this.txtLocal.TabIndex = 142;
			// 
			// btImprimir
			// 
			this.btImprimir.Location = new System.Drawing.Point(184, 128);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(64, 23);
			this.btImprimir.TabIndex = 143;
			this.btImprimir.Text = "Mostrar";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// IngresoContratosAmotiza
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(624, 165);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.txtLocal);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.txtContrato);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.btRecalculo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btMostrar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbTipoServicio);
			this.Controls.Add(this.cmbFechaAmortizacion);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbFechaP);
			this.Controls.Add(this.ultraLabel14);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "IngresoContratosAmotiza";
			this.Text = "Tabla de Amortizacion de Contratos";
			this.Load += new System.EventHandler(this.IngresoContratosAmotiza_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaAmortizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private void IngresoContratosAmotiza_Load(object sender, System.EventArgs e)
		{
			cmbFechaP.Value = DateTime.Today;
			cmbFechaAmortizacion.Value = DateTime.Today;

			#region llena combos
			string stSelect ="select idTipoServicio,Codigo from dbo.ContratosTipoServicio";
			cmbTipoServicio.DataSource = Funcion.dvProcedimiento(cdsContratos,stSelect);			
			cmbTipoServicio.Rows.Band.Columns["idTipoServicio"].Hidden = true;
			cmbTipoServicio.DisplayMember = "Codigo";
			cmbTipoServicio.ValueMember = "idTipoServicio";	
			#endregion llena combos
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "TablaAmortizacion.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "TablaAmortizacion.xls");
		}

		
		private void cdsContratos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsContratos.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirNumero);
		}
		private void btMostrar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbTipoServicio,"");
			errorProvider.SetError(btBuscar,"");

			if (cmbTipoServicio.Value == DBNull.Value ||cmbTipoServicio.Value == null || cmbTipoServicio.Value.ToString().Length <=0) 
			{
				errorProvider.SetError(cmbTipoServicio,"Seleccione Tipo...");
				return;
			}																																																													       																																																														 
			DateTime dtProceso = (DateTime) cmbFechaP.Value;
			DateTime dtAmortizacion = (DateTime) cmbFechaAmortizacion.Value;		
			string stSelect =string.Format("Exec P_ContratosTablaAmortizacion '{0}','{1}',{2},{3},1,1 ",
				dtProceso.ToString("yyyyMMdd"),dtAmortizacion.ToString("yyyyMMdd"),cmbTipoServicio.Value, 0);
			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsContratos,stSelect);		
		}
		int idContrato = 0;
		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(btGenerar, "");	
						using (IngresoContratosBuscar	miBusqueda = new IngresoContratosBuscar())
						{
							if (miBusqueda.ShowDialog() == DialogResult.OK)
							{
								try
								{						
									idContrato  = miBusqueda.idBusca;
									if (idContrato >0)
									{
										string stSelect = "Select count(*) from contratos where idpadre = "+idContrato;
										int cuenta = Funcion.iEscalarSQL(cdsContratos,stSelect);
										if (cuenta > 0)
										{
											stSelect = "Select top 1* from contratos where idpadre = "+idContrato+ "order by idcontratos desc ";
											idContrato = Funcion.iEscalarSQL(cdsContratos,stSelect);
										}					
										stSelect = "select isnull(codigo,'') from contratos where idcontratos = "+idContrato;
										txtContrato.Text = Funcion.sEscalarSQL(cdsContratos,stSelect);
										
										stSelect = "select cl.nombre from contratos c "+
											" inner join cliente as cl on cl.idcliente = c.idcliente "+
											" where idcontratos = "+idContrato;
										txtCliente.Text = Funcion.sEscalarSQL(cdsContratos,stSelect);
										
										stSelect = "select cl.nombre from contratos c "+
											" inner join ContratosLocales as cl on cl.idLocales = c.idLocales "+
											" where idcontratos = "+idContrato;			
										txtLocal.Text = Funcion.sEscalarSQL(cdsContratos,stSelect);
									}						
								}
								catch{}
							}
						}
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
		 GeneraYRegenra(0);
		}
		private void btRecalculo_Click(object sender, System.EventArgs e)
		{
		GeneraYRegenra(1);
		}

		private void GeneraYRegenra(int regenera)
		{
			this.Cursor = Cursors.WaitCursor;

			errorProvider.SetError(cmbTipoServicio,"");
			errorProvider.SetError(btBuscar,"");

			if (cmbTipoServicio.Value == DBNull.Value ||cmbTipoServicio.Value == null || cmbTipoServicio.Value.ToString().Length <=0) 
			{
				errorProvider.SetError(cmbTipoServicio,"Seleccione Tipo...");
				return;
			}	
			if(idContrato<=0)				
			{
				errorProvider.SetError(btBuscar,"Seleccione el Contrato a amortizar...");
				return;
			}	
																																												       																																																														 
			DateTime dtProceso = (DateTime) cmbFechaP.Value;
			DateTime dtAmortizacion = (DateTime) cmbFechaAmortizacion.Value;		
			string stSelect =string.Format("Exec P_ContratosTablaAmortizacion '{0}','{1}',{2},{3},0,{4} ",
				dtProceso.ToString("yyyyMMdd"),dtAmortizacion.ToString("yyyyMMdd"),cmbTipoServicio.Value,idContrato,regenera);
			//ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsContratos,stSelect);	
			string Mess = Funcion.sEscalarSQL(cdsContratos,stSelect);
			MessageBox.Show(Mess,"Informacion",MessageBoxButtons.OK,MessageBoxIcon.Information);
			//			ultraGrid1.Rows.Band.Columns["CantEjecutada"].Hidden = true;
			this.Cursor = Cursors.Default;

		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbTipoServicio,"");

			if (cmbTipoServicio.Value == DBNull.Value ||cmbTipoServicio.Value == null || cmbTipoServicio.Value.ToString().Length <=0) 
			{
				errorProvider.SetError(cmbTipoServicio,"Seleccione Tipo...");
				return;
			}	

			this.Cursor = Cursors.WaitCursor;
			string stFiltro = "{TmpContratoAmortizacion.idTipoServicio} = " + cmbTipoServicio.Value;
			string stReporte = "ContratosAmortizacion.rpt";			
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			
			this.Cursor = Cursors.Default;
		}
	}


}
