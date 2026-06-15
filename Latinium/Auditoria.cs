using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Auditoria.
	/// </summary>
	public class Auditoria : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private C1.Data.C1DataSet cdsAuditoria;
		private C1.Data.C1DataView cdvTipo;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private System.Windows.Forms.CheckBox chkPersonal;
		private System.Windows.Forms.CheckBox chkEstado;
		private System.Windows.Forms.CheckBox chkTipo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCampo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Auditoria()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AuditaEstado", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAuditoriaEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AuditaTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAuditaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdsAuditoria = new C1.Data.C1DataSet();
			this.cdvTipo = new C1.Data.C1DataView();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkPersonal = new System.Windows.Forms.CheckBox();
			this.chkEstado = new System.Windows.Forms.CheckBox();
			this.chkTipo = new System.Windows.Forms.CheckBox();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCampo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAuditoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCampo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDesde.Appearance = appearance1;
			this.cmbDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(128, 60);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(112, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 0;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(56, 60);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Desde:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(256, 60);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Hasta:";
			// 
			// cmbHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbHasta.Appearance = appearance2;
			this.cmbHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(328, 60);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(112, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 2;
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(64, 241);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(72, 26);
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(328, 241);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 26);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// cdsAuditoria
			// 
			this.cdsAuditoria.BindingContextCtrl = this;
			this.cdsAuditoria.DataLibrary = "LibPersonal";
			this.cdsAuditoria.DataLibraryUrl = "";
			this.cdsAuditoria.DataSetDef = "Auditoria";
			this.cdsAuditoria.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAuditoria.Name = "cdsAuditoria";
			this.cdsAuditoria.SchemaClassName = "LibPersonal.DataClass";
			this.cdsAuditoria.SchemaDef = null;
			this.cdsAuditoria.BeforeFill += new C1.Data.FillEventHandler(this.cdsAuditoria_BeforeFill);
			// 
			// cdvTipo
			// 
			this.cdvTipo.BindingContextCtrl = this;
			this.cdvTipo.DataSet = this.cdsAuditoria;
			this.cdvTipo.Sort = "Tipo";
			this.cdvTipo.TableName = "";
			this.cdvTipo.TableViewName = "AuditaTipo";
			// 
			// cmbEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance3;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataMember = "AuditaEstado";
			this.cmbEstado.DataSource = this.cdsAuditoria;
			this.cmbEstado.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 109;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(144, 155);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(128, 21);
			this.cmbEstado.TabIndex = 8;
			this.cmbEstado.ValueMember = "idAuditoriaEstado";
			// 
			// cmbTipo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance4;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.cdvTipo;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 109;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(144, 190);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(128, 21);
			this.cmbTipo.TabIndex = 9;
			this.cmbTipo.ValueMember = "idAuditaTipo";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkPersonal
			// 
			this.chkPersonal.Checked = true;
			this.chkPersonal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPersonal.Location = new System.Drawing.Point(24, 121);
			this.chkPersonal.Name = "chkPersonal";
			this.chkPersonal.Size = new System.Drawing.Size(104, 25);
			this.chkPersonal.TabIndex = 12;
			this.chkPersonal.Text = "Todo Personal";
			this.chkPersonal.CheckedChanged += new System.EventHandler(this.chkPersonal_CheckedChanged);
			// 
			// chkEstado
			// 
			this.chkEstado.Checked = true;
			this.chkEstado.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEstado.Location = new System.Drawing.Point(24, 155);
			this.chkEstado.Name = "chkEstado";
			this.chkEstado.Size = new System.Drawing.Size(104, 26);
			this.chkEstado.TabIndex = 13;
			this.chkEstado.Text = "Todo Estado";
			this.chkEstado.CheckedChanged += new System.EventHandler(this.chkEstado_CheckedChanged);
			// 
			// chkTipo
			// 
			this.chkTipo.Checked = true;
			this.chkTipo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTipo.Location = new System.Drawing.Point(24, 190);
			this.chkTipo.Name = "chkTipo";
			this.chkTipo.Size = new System.Drawing.Size(104, 25);
			this.chkTipo.TabIndex = 14;
			this.chkTipo.Text = "Todo Tipo";
			this.chkTipo.CheckedChanged += new System.EventHandler(this.chkTipo_CheckedChanged);
			// 
			// txtNumero
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance5;
			this.txtNumero.Location = new System.Drawing.Point(344, 190);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 15;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(288, 190);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(44, 15);
			this.ultraLabel1.TabIndex = 16;
			this.ultraLabel1.Text = "Número";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(288, 155);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(39, 15);
			this.ultraLabel2.TabIndex = 18;
			this.ultraLabel2.Text = "Campo";
			// 
			// txtCampo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCampo.Appearance = appearance6;
			this.txtCampo.Location = new System.Drawing.Point(344, 155);
			this.txtCampo.Name = "txtCampo";
			this.txtCampo.Size = new System.Drawing.Size(104, 22);
			this.txtCampo.TabIndex = 17;
			// 
			// txtUsuario
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUsuario.Appearance = appearance7;
			this.txtUsuario.Location = new System.Drawing.Point(144, 121);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(112, 22);
			this.txtUsuario.TabIndex = 19;
			// 
			// Auditoria
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(488, 294);
			this.Controls.Add(this.txtUsuario);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtCampo);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.chkTipo);
			this.Controls.Add(this.chkEstado);
			this.Controls.Add(this.chkPersonal);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Name = "Auditoria";
			this.Text = "Auditoria";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Auditoria_Closing);
			this.Load += new System.EventHandler(this.Auditoria_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAuditoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCampo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.cmbTipo, "");
			this.errorProvider.SetError(this.cmbEstado, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (!this.chkTipo.Checked || txtNumero.Text.Trim().Length > 0)
				if (this.cmbTipo.Value == null)
				{
					this.errorProvider.SetError(this.cmbTipo, "Ingrese Tipo");
					bOk = false;
				}
			if (!this.chkEstado.Checked)
				if (this.cmbEstado.Value == null)
				{
					this.errorProvider.SetError(this.cmbEstado, "Ingrese Estado");
					bOk = false;
				}

			return bOk;
		}

		private string Filtro()
		{
			string stFiltro;
			DateTime dtFechaIni = (DateTime) this.cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) this.cmbHasta.Value;
			stFiltro = "{Auditoria.Fecha} >= #" + 
				dtFechaIni.ToString("MM/dd/yyyy") + "# And " 
				+ " {Auditoria.Fecha} < #"
				+ dtFechaFin.AddDays(1).ToString("MM/dd/yyyy") + "#";
			if (!this.chkPersonal.Checked)
				stFiltro += " And {Auditoria.Usuario} = '" + txtUsuario.Text.Trim() + "'";
			
			if (!this.chkTipo.Checked)
				stFiltro += " And {Auditoria.idTipoAudita} = " + this.cmbTipo.Value.ToString();

			if (!this.chkEstado.Checked)
				stFiltro += " And {Auditoria.idAuditaEstado} = " + this.cmbEstado.Value.ToString();

			return stFiltro;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			this.Cursor = Cursors.WaitCursor;
			string stFiltro = Filtro();
			if (txtNumero.Text.Trim().Length > 0)
			{
				string stSelect;
				string stInicio;
//				if (cmbTipo.Value == null)
//				{
//					errorProvider.SetError(cmbTipo, "Ingrese Tipo");
//					return;
//				}
				switch (int.Parse(cmbTipo.Value.ToString()))
				{
					case 23:
						stSelect = "Select Top 1 idAsiento From Asiento Where CodAsiento = '"
							+ txtNumero.Text.Trim() + "' Order by idAsiento Desc";
						stInicio = "Select count(*) From Auditoria Where Codigo = '"
							+ txtNumero.Text.Trim() + "' And Campo = 'CodAsiento'";
						break;
					default:
						stSelect = "Select Top 1 idCompra From Compra Where Numero = '"
							+ txtNumero.Text.Trim() + "' Order by idCompra Desc";
						stInicio = "Select Count(*) From Compra Where Numero = '"
							+ txtNumero.Text.Trim() + "'";
						break;
				}
				int IdCuenta = Funcion.iEscalarSQL(cdsAuditoria, stInicio, false);
				if (IdCuenta == 1)
				{
					int IdCompra = Funcion.iEscalarSQL(cdsAuditoria, stSelect, false);
					stFiltro += " And {Auditoria.idUnica} = " + IdCompra.ToString();
				}
				else if (IdCuenta > 1)
				{
					stFiltro += " And {Auditoria.Codigo} = '" + txtNumero.Text.ToString().Trim() + "'";
					MessageBox.Show("Existen " + IdCuenta.ToString() + " registros", "Información");
				}
				else
				{
					MessageBox.Show("No existe codigo", "Información");
					return;
				}
			}
			if (txtCampo.Text.Trim().Length > 0)
			{
				stFiltro += " And {Auditoria.Campo} like '" 
					+ txtCampo.Text.ToString().Trim() + "*'";
			}
			Reporte miReporte = new Reporte("Auditoria.rpt", stFiltro);
			miReporte.MdiParent = this.MdiParent;
			this.Cursor = Cursors.Default;
			miReporte.Show();
		}

		private void chkEstado_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbEstado.Enabled = !this.chkEstado.Checked;
			if (!this.chkEstado.Checked) this.cmbEstado.Select();
		}

		private void chkTipo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTipo.Enabled = !this.chkTipo.Checked;
			if (!this.chkTipo.Checked) this.cmbTipo.Select();
		}

		private void chkPersonal_CheckedChanged(object sender, System.EventArgs e)
		{
		}

		private void Auditoria_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Auditoria'";
			Funcion.EjecutaSQL(cdsAuditoria, stAudita, true);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsAuditoria_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsAuditoria.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Auditoria_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Auditoria'";
			Funcion.EjecutaSQL(cdsAuditoria, stAudita, true);
		}
	}
}
