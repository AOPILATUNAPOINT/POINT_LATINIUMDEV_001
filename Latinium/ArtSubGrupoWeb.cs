using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ArtSubGrupoWeb.
	/// </summary>
	public class ArtSubGrupoWeb : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox grpGrupoWeb;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.Misc.UltraButton btnSubGrupoWebNuevo;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdSubGrupoWeb;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoWeb;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbSubGrupoWeb;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public static int idWEB_Grupos = 0;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsGrupoWeb;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSubGrupoWeb;
		public static int idWEB_Subgrupos = 0;

		public ArtSubGrupoWeb(int IdWEB_Grupos, int IdWEB_Subgrupos)
		{
			idWEB_Grupos = IdWEB_Grupos;
			idWEB_Subgrupos = IdWEB_Subgrupos;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idWEB_Grupos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idWEB_Grupos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idWEB_Subgrupos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idWEB_Subgrupos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ArtSubGrupoWeb));
			this.grpGrupoWeb = new System.Windows.Forms.GroupBox();
			this.cmbGrupoWeb = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsGrupoWeb = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.txtIdSubGrupoWeb = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnSubGrupoWebNuevo = new Infragistics.Win.Misc.UltraButton();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbSubGrupoWeb = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsSubGrupoWeb = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.grpGrupoWeb.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsGrupoWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSubGrupoWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSubGrupoWeb)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// grpGrupoWeb
			// 
			this.grpGrupoWeb.Controls.Add(this.cmbGrupoWeb);
			this.grpGrupoWeb.Controls.Add(this.label1);
			this.grpGrupoWeb.Controls.Add(this.txtIdSubGrupoWeb);
			this.grpGrupoWeb.Controls.Add(this.btnSubGrupoWebNuevo);
			this.grpGrupoWeb.Controls.Add(this.label9);
			this.grpGrupoWeb.Controls.Add(this.cmbSubGrupoWeb);
			this.grpGrupoWeb.Location = new System.Drawing.Point(16, 64);
			this.grpGrupoWeb.Name = "grpGrupoWeb";
			this.grpGrupoWeb.Size = new System.Drawing.Size(432, 136);
			this.grpGrupoWeb.TabIndex = 198;
			this.grpGrupoWeb.TabStop = false;
			this.grpGrupoWeb.Text = "SubGrupo Web";
			this.grpGrupoWeb.Enter += new System.EventHandler(this.grpGrupoWeb_Enter);
			// 
			// cmbGrupoWeb
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupoWeb.Appearance = appearance1;
			this.cmbGrupoWeb.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoWeb.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoWeb.DataSource = this.udsGrupoWeb;
			this.cmbGrupoWeb.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 93;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 298;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbGrupoWeb.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbGrupoWeb.DisplayMember = "Descripcion";
			this.cmbGrupoWeb.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoWeb.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupoWeb.Enabled = false;
			this.cmbGrupoWeb.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoWeb.Location = new System.Drawing.Point(80, 32);
			this.cmbGrupoWeb.Name = "cmbGrupoWeb";
			this.cmbGrupoWeb.Size = new System.Drawing.Size(320, 25);
			this.cmbGrupoWeb.TabIndex = 203;
			this.cmbGrupoWeb.ValueMember = "idWEB_Grupos";
			// 
			// udsGrupoWeb
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsGrupoWeb.Band.Columns.AddRange(new object[] {
																														ultraDataColumn1,
																														ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 18);
			this.label1.TabIndex = 202;
			this.label1.Text = "Grupo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdSubGrupoWeb
			// 
			this.txtIdSubGrupoWeb.Enabled = false;
			this.txtIdSubGrupoWeb.Location = new System.Drawing.Point(408, 80);
			this.txtIdSubGrupoWeb.Name = "txtIdSubGrupoWeb";
			this.txtIdSubGrupoWeb.Size = new System.Drawing.Size(10, 24);
			this.txtIdSubGrupoWeb.TabIndex = 197;
			this.txtIdSubGrupoWeb.Visible = false;
			// 
			// btnSubGrupoWebNuevo
			// 
			this.btnSubGrupoWebNuevo.Location = new System.Drawing.Point(344, 80);
			this.btnSubGrupoWebNuevo.Name = "btnSubGrupoWebNuevo";
			this.btnSubGrupoWebNuevo.Size = new System.Drawing.Size(56, 28);
			this.btnSubGrupoWebNuevo.TabIndex = 176;
			this.btnSubGrupoWebNuevo.Text = "Nuevo";
			this.btnSubGrupoWebNuevo.Click += new System.EventHandler(this.btnSubGrupoWebNuevo_Click);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(16, 88);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(31, 18);
			this.label9.TabIndex = 175;
			this.label9.Text = "Tipo";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSubGrupoWeb
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubGrupoWeb.Appearance = appearance2;
			this.cmbSubGrupoWeb.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubGrupoWeb.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubGrupoWeb.DataSource = this.udsSubGrupoWeb;
			this.cmbSubGrupoWeb.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.Caption = "idWEB_SubGrupos";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 86;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 226;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbSubGrupoWeb.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSubGrupoWeb.DisplayMember = "Descripcion";
			this.cmbSubGrupoWeb.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubGrupoWeb.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubGrupoWeb.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubGrupoWeb.Location = new System.Drawing.Point(80, 80);
			this.cmbSubGrupoWeb.Name = "cmbSubGrupoWeb";
			this.cmbSubGrupoWeb.Size = new System.Drawing.Size(248, 25);
			this.cmbSubGrupoWeb.TabIndex = 174;
			this.cmbSubGrupoWeb.ValueMember = "idWEB_Subgrupos";
			// 
			// udsSubGrupoWeb
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsSubGrupoWeb.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn3,
																															 ultraDataColumn4});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(16, 216);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(92, 27);
			this.btnGuardar.TabIndex = 199;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(120, 216);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(86, 27);
			this.btnCancelar.TabIndex = 200;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Location = new System.Drawing.Point(16, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(432, 56);
			this.groupBox2.TabIndex = 201;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Indicaciones";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(348, 18);
			this.label2.TabIndex = 172;
			this.label2.Text = "* Este módulo permite selecciolnar el subgrupo en la web.";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// ArtSubGrupoWeb
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(464, 256);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.grpGrupoWeb);
			this.Name = "ArtSubGrupoWeb";
			this.Text = "SubGrupo Web";
			this.Load += new System.EventHandler(this.ArtSubGrupoWeb_Load);
			this.grpGrupoWeb.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsGrupoWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSubGrupoWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubGrupoWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSubGrupoWeb)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ArtSubGrupoWeb_Load(object sender, System.EventArgs e)
		{
			this.cmbGrupoWeb.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "EXEC WEB_GruposCon");
			this.cmbGrupoWeb.Value = idWEB_Grupos;
			string squery = string.Format("EXEC WEB_SubgruposXIdWEB_Grupos {0}", idWEB_Grupos);
			this.cmbSubGrupoWeb.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
			this.cmbSubGrupoWeb.Value = idWEB_Subgrupos;
		}

		private void grpGrupoWeb_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void Aceptar()
		{
			if (this.cmbGrupoWeb.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un grupo web", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbGrupoWeb.Focus();
				return;
			}
			else if (this.cmbSubGrupoWeb.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un subgrupo web", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbGrupoWeb.Focus();
				return;
			}
			else 
			{
				if (Funcion.ValidaBloqueoERMEnTransacciones((int)this.cmbSubGrupoWeb.Value, cdsSeteoF)) return;
				if (Funcion.ValidaBloqueo((int)this.cmbSubGrupoWeb.Value, 1, cdsSeteoF)) return;
				this.DialogResult = DialogResult.OK;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnSubGrupoWebNuevo_Click(object sender, System.EventArgs e)
		{
			int getIdGrupoWeb = 0;
			if (this.cmbGrupoWeb.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un grupo web", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbGrupoWeb.Focus();
				return;
			}
			getIdGrupoWeb = (int)this.cmbGrupoWeb.Value;

			using (ArtSubGrupoWebNuevo miSubGruposWeb = new ArtSubGrupoWebNuevo(getIdGrupoWeb))
			{
				if (DialogResult.OK == miSubGruposWeb.ShowDialog())
				{
					string squery = string.Format("EXEC WEB_SubgruposXIdWEB_Grupos {0}", idWEB_Grupos);
					this.cmbSubGrupoWeb.DataSource = Funcion.dvProcedimiento(cdsSeteoF, squery);
					this.cmbSubGrupoWeb.Value = idWEB_Subgrupos;
				}
			}

		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

	}
}
