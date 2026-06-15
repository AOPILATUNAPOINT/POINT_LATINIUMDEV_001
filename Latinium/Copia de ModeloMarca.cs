using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for ModeloMarca.
	/// </summary>
	public class ModeloMarca : DevExpress.XtraEditors.XtraForm
	{
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbModelo;
		private C1.Data.C1DataSet cdsGrupoSubgrupo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdMarcaModelo;

		C1.Data.C1DataRow drMarcaModelo;

		public ModeloMarca()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloMarca", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloModelo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsGrupoSubgrupo = new C1.Data.C1DataSet();
			this.txtIdMarcaModelo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbModelo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoSubgrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMarcaModelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).BeginInit();
			this.SuspendLayout();
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "ModeloMarca";
			this.barraDato1.DataNombreId = "idMarcaModelo";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsGrupoSubgrupo;
			this.barraDato1.DataTabla = "ModeloMarca";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(74, 72);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(254, 26);
			this.barraDato1.TabIndex = 231;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsGrupoSubgrupo
			// 
			this.cdsGrupoSubgrupo.BindingContextCtrl = this;
			this.cdsGrupoSubgrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupoSubgrupo.DataLibraryUrl = "";
			this.cdsGrupoSubgrupo.DataSetDef = "dsGruposMarcas";
			this.cdsGrupoSubgrupo.EnforceConstraints = false;
			this.cdsGrupoSubgrupo.FillOnRequest = false;
			this.cdsGrupoSubgrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoSubgrupo.Name = "cdsGrupoSubgrupo";
			this.cdsGrupoSubgrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoSubgrupo.SchemaDef = null;
			this.cdsGrupoSubgrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupoSubgrupo_BeforeFill);
			// 
			// txtIdMarcaModelo
			// 
			this.txtIdMarcaModelo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdMarcaModelo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoSubgrupo, "ModeloMarca.idMarcaModelo"));
			this.txtIdMarcaModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdMarcaModelo.Location = new System.Drawing.Point(16, 72);
			this.txtIdMarcaModelo.Name = "txtIdMarcaModelo";
			this.txtIdMarcaModelo.PromptChar = ' ';
			this.txtIdMarcaModelo.Size = new System.Drawing.Size(24, 22);
			this.txtIdMarcaModelo.TabIndex = 230;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 229;
			this.label1.Text = "Modelo";
			// 
			// cmbMarca
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance1;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoSubgrupo, "ModeloMarca.idMarca"));
			this.cmbMarca.DataMember = "ArticuloMarca";
			this.cmbMarca.DataSource = this.cdsArticuloTabla;
			appearance2.BackColor = System.Drawing.Color.White;
			this.cmbMarca.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 315;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.cmbMarca.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbMarca.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.cmbMarca.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbMarca.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbMarca.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			this.cmbMarca.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbMarca.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbMarca.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.cmbMarca.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(80, 8);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(315, 21);
			this.cmbMarca.TabIndex = 226;
			this.cmbMarca.ValueMember = "idMarca";
			// 
			// cdsArticuloTabla
			// 
			this.cdsArticuloTabla.BindingContextCtrl = this;
			this.cdsArticuloTabla.CaseSensitive = false;
			this.cdsArticuloTabla.DataLibrary = "LibFacturacion";
			this.cdsArticuloTabla.DataLibraryUrl = "";
			this.cdsArticuloTabla.DataSetDef = "dsArticuloTabla";
			this.cdsArticuloTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticuloTabla.Name = "cdsArticuloTabla";
			this.cdsArticuloTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticuloTabla.SchemaDef = null;
			this.cdsArticuloTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticuloTabla_BeforeFill);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(8, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(34, 16);
			this.label9.TabIndex = 227;
			this.label9.Text = "Marca";
			// 
			// cmbModelo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbModelo.Appearance = appearance7;
			this.cmbModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbModelo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbModelo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoSubgrupo, "ModeloMarca.idModelo"));
			this.cmbModelo.DataMember = "ArticuloModelo";
			this.cmbModelo.DataSource = this.cdsArticuloTabla;
			appearance8.BackColor = System.Drawing.Color.White;
			this.cmbModelo.DisplayLayout.Appearance = appearance8;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 315;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbModelo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.cmbModelo.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbModelo.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.cmbModelo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbModelo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbModelo.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			this.cmbModelo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbModelo.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbModelo.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.cmbModelo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbModelo.DisplayMember = "Modelo";
			this.cmbModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbModelo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbModelo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbModelo.Location = new System.Drawing.Point(80, 40);
			this.cmbModelo.Name = "cmbModelo";
			this.cmbModelo.Size = new System.Drawing.Size(315, 21);
			this.cmbModelo.TabIndex = 232;
			this.cmbModelo.ValueMember = "idModelo";
			// 
			// ModeloMarca
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(402, 104);
			this.Controls.Add(this.cmbModelo);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtIdMarcaModelo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.cmbMarca);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ModeloMarca";
			this.Load += new System.EventHandler(this.ModeloMarca_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoSubgrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMarcaModelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drMarcaModelo = cdsGrupoSubgrupo.TableViews["ModeloMarca"].Rows[0];
			}
			catch{};
		}

		private void cdsArticuloTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticuloTabla.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsGrupoSubgrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsGrupoSubgrupo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ModeloMarca_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.PosUltima();
			barraDato1.HabilitaControles(false);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.barraDato1.GrabaRegistro();
		}
	}
}

