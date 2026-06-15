using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for Alarma.
	/// </summary>
	public class Alarma : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.DateEdit cmbDesde;
		private DevExpress.XtraEditors.DateEdit cmbHasta;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.MemoEdit txtNota;
		private DevExpress.XtraEditors.CheckEdit chkActivar;
		private C1.Data.C1DataSet cdsAlarma;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraGrid.Columns.GridColumn colDesde;
		private DevExpress.XtraGrid.Columns.GridColumn colHasta;
		private DevExpress.XtraGrid.Columns.GridColumn colMensaje;
		private DevExpress.XtraGrid.Columns.GridColumn colActivar;
		private DevExpress.XtraGrid.Columns.GridColumn colUsuario;
		private DevExpress.XtraEditors.Repository.RepositoryItemDateEdit repositoryItemDateEdit1;
		private DevExpress.XtraEditors.TextEdit textEdit1;
		private DevExpress.XtraEditors.LabelControl labelControl3;
		private C1.Data.C1DataSet cdsAlarmas;
		private DevExpress.XtraEditors.SimpleButton btActivos;
		private DevExpress.XtraEditors.LabelControl labelControl4;
		private DevExpress.XtraEditors.LabelControl labelControl5;
		private DevExpress.XtraEditors.DateEdit dateEdit1;
		private DevExpress.XtraEditors.DateEdit dateEdit2;
		private System.Windows.Forms.TextBox TxtIdAlarma1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Alarma()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

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
			this.cmbDesde = new DevExpress.XtraEditors.DateEdit();
			this.cdsAlarma = new C1.Data.C1DataSet();
			this.cmbHasta = new DevExpress.XtraEditors.DateEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.txtNota = new DevExpress.XtraEditors.MemoEdit();
			this.chkActivar = new DevExpress.XtraEditors.CheckEdit();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.cdsAlarmas = new C1.Data.C1DataSet();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.colDesde = new DevExpress.XtraGrid.Columns.GridColumn();
			this.repositoryItemDateEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();
			this.colHasta = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colMensaje = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colActivar = new DevExpress.XtraGrid.Columns.GridColumn();
			this.colUsuario = new DevExpress.XtraGrid.Columns.GridColumn();
			this.textEdit1 = new DevExpress.XtraEditors.TextEdit();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.btActivos = new DevExpress.XtraEditors.SimpleButton();
			this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
			this.dateEdit1 = new DevExpress.XtraEditors.DateEdit();
			this.dateEdit2 = new DevExpress.XtraEditors.DateEdit();
			this.TxtIdAlarma1 = new System.Windows.Forms.TextBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkActivar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarmas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit2.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			this.cmbDesde.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "Alarma.Desde"));
			this.cmbDesde.EditValue = null;
			this.cmbDesde.Location = new System.Drawing.Point(80, 16);
			this.cmbDesde.Name = "cmbDesde";
			// 
			// cmbDesde.Properties
			// 
			this.cmbDesde.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbDesde.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbDesde.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbDesde.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.cmbDesde.Size = new System.Drawing.Size(88, 20);
			this.cmbDesde.TabIndex = 0;
			// 
			// cdsAlarma
			// 
			this.cdsAlarma.BindingContextCtrl = this;
			this.cdsAlarma.DataLibrary = "LibFacturacion";
			this.cdsAlarma.DataLibraryUrl = "";
			this.cdsAlarma.DataSetDef = "dsAgenda";
			this.cdsAlarma.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAlarma.Name = "cdsAlarma";
			this.cdsAlarma.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsAlarma.SchemaDef = null;
			this.cdsAlarma.BeforeFill += new C1.Data.FillEventHandler(this.cdsAgenda_BeforeFill);
			// 
			// cmbHasta
			// 
			this.cmbHasta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "Alarma.Hasta"));
			this.cmbHasta.EditValue = null;
			this.cmbHasta.Location = new System.Drawing.Point(240, 16);
			this.cmbHasta.Name = "cmbHasta";
			// 
			// cmbHasta.Properties
			// 
			this.cmbHasta.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbHasta.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy HH:mm";
			this.cmbHasta.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbHasta.Properties.EditFormat.FormatString = "dd/MMM/yyyy HH:mm";
			this.cmbHasta.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.cmbHasta.Properties.Mask.EditMask = "dd/MMM/yyyy HH:mm";
			this.cmbHasta.Size = new System.Drawing.Size(88, 20);
			this.cmbHasta.TabIndex = 1;
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(24, 16);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(40, 14);
			this.labelControl1.TabIndex = 2;
			this.labelControl1.Text = "Desde:";
			// 
			// labelControl2
			// 
			this.labelControl2.Location = new System.Drawing.Point(184, 16);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(32, 14);
			this.labelControl2.TabIndex = 3;
			this.labelControl2.Text = "Hasta:";
			// 
			// txtNota
			// 
			this.txtNota.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.txtNota.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "Alarma.Mensaje"));
			this.txtNota.Location = new System.Drawing.Point(16, 112);
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(336, 216);
			this.txtNota.TabIndex = 4;
			// 
			// chkActivar
			// 
			this.chkActivar.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "Alarma.Activar"));
			this.chkActivar.Location = new System.Drawing.Point(24, 80);
			this.chkActivar.Name = "chkActivar";
			// 
			// chkActivar.Properties
			// 
			this.chkActivar.Properties.Caption = "Activar";
			this.chkActivar.Size = new System.Drawing.Size(56, 18);
			this.chkActivar.TabIndex = 5;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Alarma";
			this.barraDato1.DataNombreId = "idAlarma";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsAlarma;
			this.barraDato1.DataTabla = "Alarma";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 344);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(288, 26);
			this.barraDato1.TabIndex = 11;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataMember = "Alarma";
			this.gridControl1.DataSource = this.cdsAlarmas;
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(368, 16);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemDateEdit1});
			this.gridControl1.Size = new System.Drawing.Size(424, 352);
			this.gridControl1.TabIndex = 13;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// cdsAlarmas
			// 
			this.cdsAlarmas.BindingContextCtrl = this;
			this.cdsAlarmas.DataLibrary = "LibFacturacion";
			this.cdsAlarmas.DataLibraryUrl = "";
			this.cdsAlarmas.DataSetDef = "dsAgenda";
			this.cdsAlarmas.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAlarmas.Name = "cdsAlarmas";
			this.cdsAlarmas.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsAlarmas.SchemaDef = null;
			this.cdsAlarmas.CurrentRowChanged += new C1.Data.CurrentRowChangedEventHandler(this.cdsAlarmas_CurrentRowChanged);
			this.cdsAlarmas.BeforeFill += new C1.Data.FillEventHandler(this.cdsAlarmas_BeforeFill);
			// 
			// gridView1
			// 
			this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
																																										 this.colDesde,
																																										 this.colHasta,
																																										 this.colMensaje,
																																										 this.colActivar,
																																										 this.colUsuario});
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			// 
			// colDesde
			// 
			this.colDesde.Caption = "Desde";
			this.colDesde.ColumnEdit = this.repositoryItemDateEdit1;
			this.colDesde.FieldName = "Desde";
			this.colDesde.Name = "colDesde";
			this.colDesde.OptionsColumn.AllowEdit = false;
			this.colDesde.OptionsColumn.FixedWidth = true;
			this.colDesde.Visible = true;
			this.colDesde.VisibleIndex = 0;
			this.colDesde.Width = 114;
			// 
			// repositoryItemDateEdit1
			// 
			this.repositoryItemDateEdit1.AutoHeight = false;
			this.repositoryItemDateEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemDateEdit1.DisplayFormat.FormatString = "dd/MMM/yyyy HH:mm";
			this.repositoryItemDateEdit1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.repositoryItemDateEdit1.EditFormat.FormatString = "dd/MMM/yyyy HH:mm";
			this.repositoryItemDateEdit1.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.repositoryItemDateEdit1.Mask.EditMask = "dd/MMM/yyyy HH:mm";
			this.repositoryItemDateEdit1.Name = "repositoryItemDateEdit1";
			// 
			// colHasta
			// 
			this.colHasta.Caption = "Hasta";
			this.colHasta.ColumnEdit = this.repositoryItemDateEdit1;
			this.colHasta.FieldName = "Hasta";
			this.colHasta.Name = "colHasta";
			this.colHasta.OptionsColumn.AllowEdit = false;
			this.colHasta.OptionsColumn.FixedWidth = true;
			this.colHasta.Visible = true;
			this.colHasta.VisibleIndex = 1;
			this.colHasta.Width = 109;
			// 
			// colMensaje
			// 
			this.colMensaje.Caption = "Mensaje";
			this.colMensaje.FieldName = "Mensaje";
			this.colMensaje.Name = "colMensaje";
			this.colMensaje.OptionsColumn.AllowEdit = false;
			this.colMensaje.Visible = true;
			this.colMensaje.VisibleIndex = 2;
			this.colMensaje.Width = 56;
			// 
			// colActivar
			// 
			this.colActivar.Caption = "Activar";
			this.colActivar.FieldName = "Activar";
			this.colActivar.Name = "colActivar";
			this.colActivar.OptionsColumn.AllowEdit = false;
			this.colActivar.Visible = true;
			this.colActivar.VisibleIndex = 3;
			this.colActivar.Width = 56;
			// 
			// colUsuario
			// 
			this.colUsuario.Caption = "Usuario";
			this.colUsuario.FieldName = "Usuario";
			this.colUsuario.Name = "colUsuario";
			this.colUsuario.OptionsColumn.AllowEdit = false;
			this.colUsuario.Visible = true;
			this.colUsuario.VisibleIndex = 4;
			this.colUsuario.Width = 68;
			// 
			// textEdit1
			// 
			this.textEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "Alarma.Usuario"));
			this.textEdit1.Location = new System.Drawing.Point(176, 80);
			this.textEdit1.Name = "textEdit1";
			this.textEdit1.Size = new System.Drawing.Size(152, 20);
			this.textEdit1.TabIndex = 14;
			// 
			// labelControl3
			// 
			this.labelControl3.Location = new System.Drawing.Point(120, 80);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(48, 16);
			this.labelControl3.TabIndex = 15;
			this.labelControl3.Text = "Usuario";
			// 
			// btActivos
			// 
			this.btActivos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btActivos.Location = new System.Drawing.Point(312, 344);
			this.btActivos.Name = "btActivos";
			this.btActivos.Size = new System.Drawing.Size(48, 24);
			this.btActivos.TabIndex = 16;
			this.btActivos.Text = "Activo";
			this.btActivos.Click += new System.EventHandler(this.btActivos_Click);
			// 
			// labelControl4
			// 
			this.labelControl4.Location = new System.Drawing.Point(184, 48);
			this.labelControl4.Name = "labelControl4";
			this.labelControl4.Size = new System.Drawing.Size(32, 14);
			this.labelControl4.TabIndex = 20;
			this.labelControl4.Text = "Final:";
			// 
			// labelControl5
			// 
			this.labelControl5.Location = new System.Drawing.Point(24, 48);
			this.labelControl5.Name = "labelControl5";
			this.labelControl5.Size = new System.Drawing.Size(48, 14);
			this.labelControl5.TabIndex = 19;
			this.labelControl5.Text = "Inicio:";
			// 
			// dateEdit1
			// 
			this.dateEdit1.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "Alarma.HoraHasta"));
			this.dateEdit1.EditValue = null;
			this.dateEdit1.Location = new System.Drawing.Point(240, 48);
			this.dateEdit1.Name = "dateEdit1";
			// 
			// dateEdit1.Properties
			// 
			this.dateEdit1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dateEdit1.Properties.DisplayFormat.FormatString = "HH:mm";
			this.dateEdit1.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit1.Properties.EditFormat.FormatString = "HH:mm";
			this.dateEdit1.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit1.Properties.Mask.EditMask = "HH:mm";
			this.dateEdit1.Size = new System.Drawing.Size(88, 20);
			this.dateEdit1.TabIndex = 18;
			// 
			// dateEdit2
			// 
			this.dateEdit2.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAlarma, "Alarma.HoraDesde"));
			this.dateEdit2.EditValue = null;
			this.dateEdit2.Location = new System.Drawing.Point(80, 48);
			this.dateEdit2.Name = "dateEdit2";
			// 
			// dateEdit2.Properties
			// 
			this.dateEdit2.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dateEdit2.Properties.DisplayFormat.FormatString = "HH:mm";
			this.dateEdit2.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit2.Properties.EditFormat.FormatString = "HH:mm";
			this.dateEdit2.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit2.Properties.Mask.EditMask = "HH:mm";
			this.dateEdit2.Size = new System.Drawing.Size(88, 20);
			this.dateEdit2.TabIndex = 17;
			// 
			// TxtIdAlarma1
			// 
			this.TxtIdAlarma1.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsAlarmas, "Alarma.idAlarma"));
			this.TxtIdAlarma1.Location = new System.Drawing.Point(328, 0);
			this.TxtIdAlarma1.Name = "TxtIdAlarma1";
			this.TxtIdAlarma1.Size = new System.Drawing.Size(40, 20);
			this.TxtIdAlarma1.TabIndex = 21;
			this.TxtIdAlarma1.Text = "";
			// 
			// Alarma
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(800, 382);
			this.Controls.Add(this.TxtIdAlarma1);
			this.Controls.Add(this.labelControl4);
			this.Controls.Add(this.labelControl5);
			this.Controls.Add(this.dateEdit1);
			this.Controls.Add(this.dateEdit2);
			this.Controls.Add(this.btActivos);
			this.Controls.Add(this.labelControl3);
			this.Controls.Add(this.textEdit1);
			this.Controls.Add(this.gridControl1);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.chkActivar);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.labelControl2);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Name = "Alarma";
			this.Text = "Alarma";
			this.Load += new System.EventHandler(this.Alarma_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkActivar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarmas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit2.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsAgenda_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsAlarma.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void Alarma_Load(object sender, System.EventArgs e)
		{
			TxtIdAlarma1.Width = 0;
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();
			cmbDesde.SelectAll();
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			barraDato1.BorraRegistro();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
      barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
		}

		private void cdsAlarmas_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsAlarmas.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);

		}

		private void cdsAlarmas_CurrentRowChanged(object sender, C1.Data.CurrentRowChangedEventArgs e)
		{
//			barraDato1.IdRegistro = gridv
			barraDato1.IdRegistro = int.Parse(TxtIdAlarma1.Text);
			barraDato1.ProximoId(5);
		}

		private void btActivos_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec AlarmaActivar 'Todos', 1";
			int iCuenta = Funcion.iEscalarSQL(cdsAlarma, stExec);
			if (iCuenta == 0)
			{
				MessageBox.Show("No existen alarmas en el periodo", "Información");
				return;
			}
			stExec = "Exec AlarmaActivar 'Todos', 2";
			gridView1.Columns.Clear();
			gridControl1.DataSource = Funcion.dvProcedimiento(cdsAlarma, stExec);
		}
	}
}

