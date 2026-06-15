using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ArtSubGrupoWebNuevoCon.
	/// </summary>
	public class ArtSubGrupoWebNuevoCon : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSubGrupoWeb;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdSubGrupoWeb;
		public static int idWEB_Grupos = 0;

		public ArtSubGrupoWebNuevoCon(int IdWEB_Grupos)
		{
			idWEB_Grupos = IdWEB_Grupos;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idWEB_Subgrupos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idWEB_Subgrupos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.grdSubGrupoWeb = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsSubGrupoWeb = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdSubGrupoWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSubGrupoWeb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Location = new System.Drawing.Point(16, 440);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(90, 27);
			this.btnAceptar.TabIndex = 181;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(120, 440);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(90, 27);
			this.btnCancelar.TabIndex = 182;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Location = new System.Drawing.Point(16, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(480, 56);
			this.groupBox2.TabIndex = 180;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Indicaciones";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(388, 18);
			this.label1.TabIndex = 172;
			this.label1.Text = "* Este módulo permite seleccionar el subgrupo que será editado.";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grdSubGrupoWeb
			// 
			this.grdSubGrupoWeb.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdSubGrupoWeb.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdSubGrupoWeb.DataSource = this.udsSubGrupoWeb;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdSubGrupoWeb.DisplayLayout.Appearance = appearance1;
			this.grdSubGrupoWeb.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 107;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 459;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			ultraGridBand1.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.grdSubGrupoWeb.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdSubGrupoWeb.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdSubGrupoWeb.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 9F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSubGrupoWeb.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.grdSubGrupoWeb.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSubGrupoWeb.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSubGrupoWeb.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSubGrupoWeb.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdSubGrupoWeb.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSubGrupoWeb.Location = new System.Drawing.Point(16, 80);
			this.grdSubGrupoWeb.Name = "grdSubGrupoWeb";
			this.grdSubGrupoWeb.Size = new System.Drawing.Size(480, 344);
			this.grdSubGrupoWeb.TabIndex = 179;
			// 
			// udsSubGrupoWeb
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsSubGrupoWeb.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn1,
																															 ultraDataColumn2});
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
			// ArtSubGrupoWebNuevoCon
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(520, 488);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.grdSubGrupoWeb);
			this.Name = "ArtSubGrupoWebNuevoCon";
			this.Text = "ArtSubGrupoWebNuevoCon";
			this.Load += new System.EventHandler(this.ArtSubGrupoWebNuevoCon_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdSubGrupoWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSubGrupoWeb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ArtSubGrupoWebNuevoCon_Load(object sender, System.EventArgs e)
		{
			this.grdSubGrupoWeb.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC WEB_SubgruposXIdWEB_Grupos {0}", idWEB_Grupos));
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
						Aceptar();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
						this.DialogResult = DialogResult.No;
		}

		private void Aceptar()
		{
			if (this.grdSubGrupoWeb.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros encontrados.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);	
			}
			else if (this.grdSubGrupoWeb.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un registro.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (this.grdSubGrupoWeb.ActiveRow.Cells["idWEB_Subgrupos"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un registro.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else 
			{
				if (Funcion.ValidaBloqueoERMEnTransacciones((int)this.grdSubGrupoWeb.ActiveRow.Cells["idWEB_Subgrupos"].Value, cdsSeteoF)) return;
				if (Funcion.ValidaBloqueo((int)this.grdSubGrupoWeb.ActiveRow.Cells["idWEB_Subgrupos"].Value, 1, cdsSeteoF)) return;
				this.DialogResult = DialogResult.OK;
			}
		}

	}
}
