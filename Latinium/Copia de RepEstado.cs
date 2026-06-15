using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepEstado.
	/// </summary>
	public class RepEstado : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsReporte;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataView cdvReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private C1.Data.C1DataSet cdsSeteoC;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProyecto;
		private DevExpress.XtraEditors.CheckEdit chkSubProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbProyecto;
		private DevExpress.XtraEditors.LookUpEdit cmbSubProyecto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepEstado()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.chkProyecto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSubProyecto = new DevExpress.XtraEditors.CheckEdit();
			this.cmbProyecto = new DevExpress.XtraEditors.LookUpEdit();
			this.cmbSubProyecto = new DevExpress.XtraEditors.LookUpEdit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 64);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 8;
			this.label2.Text = "Hasta:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 7;
			this.label1.Text = "Desde:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 168);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 45;
			this.label3.Text = "Reporte";
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(64, 208);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 46;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(200, 208);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 47;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 4";
			this.cdvReporte.Sort = "Detalle";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 345;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 54;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 54;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbReporte.DisplayMember = "Detalle";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(160, 160);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(160, 21);
			this.cmbReporte.TabIndex = 48;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// cmbDesde
			// 
			this.cmbDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(160, 32);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 24;
			this.cmbDesde.Size = new System.Drawing.Size(128, 21);
			this.cmbDesde.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 49;
			this.cmbDesde.Value = new System.DateTime(2009, 7, 30, 0, 0, 0, 0);
			// 
			// cmbHasta
			// 
			this.cmbHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(160, 64);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 24;
			this.cmbHasta.Size = new System.Drawing.Size(128, 21);
			this.cmbHasta.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 50;
			this.cmbHasta.Value = new System.DateTime(2009, 7, 30, 0, 0, 0, 0);
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible = 1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// chkProyecto
			// 
			this.chkProyecto.Checked = true;
			this.chkProyecto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkProyecto.Location = new System.Drawing.Point(24, 96);
			this.chkProyecto.Name = "chkProyecto";
			this.chkProyecto.Size = new System.Drawing.Size(125, 16);
			this.chkProyecto.TabIndex = 64;
			this.chkProyecto.Text = "Todo Centro Costo";
			this.chkProyecto.CheckedChanged += new System.EventHandler(this.chkProyecto_CheckedChanged);
			// 
			// chkSubProyecto
			// 
			this.chkSubProyecto.EditValue = true;
			this.chkSubProyecto.Location = new System.Drawing.Point(24, 128);
			this.chkSubProyecto.Name = "chkSubProyecto";
			// 
			// chkSubProyecto.Properties
			// 
			this.chkSubProyecto.Properties.Caption = "Todo Proyecto";
			this.chkSubProyecto.Size = new System.Drawing.Size(128, 19);
			this.chkSubProyecto.TabIndex = 63;
			this.chkSubProyecto.CheckedChanged += new System.EventHandler(this.chkSubProyecto_CheckedChanged);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Location = new System.Drawing.Point(160, 96);
			this.cmbProyecto.Name = "cmbProyecto";
			// 
			// cmbProyecto.Properties
			// 
			this.cmbProyecto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbProyecto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbProyecto.Properties.DataSource = this.cdvProyecto;
			this.cmbProyecto.Properties.DisplayMember = "Nombre";
			this.cmbProyecto.Properties.Enabled = false;
			this.cmbProyecto.Properties.NullText = "";
			this.cmbProyecto.Properties.PopupWidth = 300;
			this.cmbProyecto.Properties.ValueMember = "idProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(160, 21);
			this.cmbProyecto.TabIndex = 62;
			// 
			// cmbSubProyecto
			// 
			this.cmbSubProyecto.Location = new System.Drawing.Point(160, 128);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			// 
			// cmbSubProyecto.Properties
			// 
			this.cmbSubProyecto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbSubProyecto.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near)});
			this.cmbSubProyecto.Properties.DataSource = this.cdvSubProyecto;
			this.cmbSubProyecto.Properties.DisplayMember = "Nombre";
			this.cmbSubProyecto.Properties.Enabled = false;
			this.cmbSubProyecto.Properties.NullText = "";
			this.cmbSubProyecto.Properties.PopupWidth = 300;
			this.cmbSubProyecto.Properties.ValueMember = "idSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(160, 21);
			this.cmbSubProyecto.TabIndex = 61;
			// 
			// RepEstado
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(344, 254);
			this.Controls.Add(this.chkProyecto);
			this.Controls.Add(this.chkSubProyecto);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbSubProyecto);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.MaximizeBox = false;
			this.Name = "RepEstado";
			this.Text = "Rep. Contabilidad NIIF";
			this.Load += new System.EventHandler(this.RepEstado_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkSubProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(cmbReporte, "");
			errorProvider.SetError(cmbProyecto, "");
			errorProvider.SetError(cmbSubProyecto, "");
			if (cmbReporte.Text.Trim().Length < 1)
			{
				errorProvider.SetError(cmbReporte, "Ingrese Reporte");
				bOk = false;
			}
			if (!chkProyecto.Checked)
			{
				if (cmbProyecto.Text.Trim().Length < 1)
				{
					errorProvider.SetError(cmbProyecto, "Ingrese Centro de Costo");
					bOk = false;
				}
			}

			if (!chkSubProyecto.Checked)
			{
				if (cmbSubProyecto.Text.Trim().Length < 1)
				{
					errorProvider.SetError(cmbSubProyecto, "Ingrese Proyecto");
					bOk = false;
				}
			}
			return bOk;
		}
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stReporte = cmbReporte.ActiveRow.Cells["Reporte"].Value.ToString();
			stReporte = stReporte.Trim();
			string stProyecto;
			if (chkProyecto.Checked) stProyecto = ", 0";
			else stProyecto = ", " + cmbProyecto.EditValue.ToString();

			if (chkSubProyecto.Checked) stProyecto += ", 0";
			else stProyecto += ", " + cmbSubProyecto.EditValue.ToString();

			int iTipo = (int) cmbReporte.ActiveRow.Cells["idGrupo"].Value;
			if (iTipo < 20 && iTipo > 10) iTipo -= 8;
			if (iTipo < 30 && iTipo > 20) iTipo -= 16;
			string stTipo = "', " + iTipo.ToString();
			if (iTipo < 10) 
			{
				stReporte = stReporte.Substring(0, stReporte.Length -1);
				string stExec = "Exec EstadoFlujo '" + dtDesde.ToString("yyyyMMdd") 
					+ "', '" + dtHasta.ToString("yyyyMMdd") + stTipo + stProyecto;
				Funcion.EjecutaSQL(cdsSeteoC, stExec, true);
			}
			else
			{
				string stExec = "Exec PresupuestoCedula '" + dtHasta.ToString("yyyyMMdd") + "'";
				Funcion.EjecutaSQL(cdsSeteoC, stExec, true);
			}
			string stTitulo = cmbReporte.Text;
			
			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "Periodo";
			discreteVal.Value = "Desde el " + dtDesde.ToString("dd/MMM/yyyy")
				+ " Hasta el " + dtHasta.ToString("dd/MMM/yyyy");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			Reporte miRep = new Reporte(stReporte);
			miRep.Titulo(stTitulo);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void RepEstado_Load(object sender, System.EventArgs e)
		{
			cdvReporte.RowFilter = "Tipo = 8";
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void chkProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
//			if (!this.chkProyecto.Checked) this.chkProyecto.Checked = true;
			this.cmbProyecto.Enabled = !this.chkProyecto.Checked;
			this.cmbProyecto.Focus();		
		}

		private void chkSubProyecto_CheckedChanged(object sender, System.EventArgs e)
		{
//			if (!this.chkSubProyecto.Checked) this.chkSubProyecto.Checked = true;
			this.cmbSubProyecto.Enabled = !this.chkSubProyecto.Checked;
			this.cmbSubProyecto.Focus();
		}
	}
}
