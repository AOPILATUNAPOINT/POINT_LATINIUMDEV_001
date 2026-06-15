using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de MarcasDeArticulo.
	/// </summary>
	public class MarcasDeArticulo : System.Windows.Forms.Form
	{
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataView cvdArticuloGrupo;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private C1.Data.C1DataSet cdsMarca;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdMarca;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		C1.Data.C1DataRow drMarca;

		public MarcasDeArticulo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentas9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.cvdArticuloGrupo = new C1.Data.C1DataView();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.cdsMarca = new C1.Data.C1DataSet();
			this.txtIdMarca = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cvdArticuloGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMarca)).BeginInit();
			this.SuspendLayout();
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "ArticuloMarca";
			this.barraDato1.DataNombreId = "idMarca";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsMarca;
			this.barraDato1.DataTabla = "ArticuloMarca";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(72, 90);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 104;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtCodigo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarca, "ArticuloMarca.Marca"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(120, 50);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(304, 21);
			this.txtCodigo.TabIndex = 102;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(24, 50);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 103;
			this.label1.Text = "Marca";
			// 
			// cmbGrupo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance2;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarca, "ArticuloMarca.idGrupoArticulo"));
			this.cmbGrupo.DataSource = this.cvdArticuloGrupo;
			appearance3.BackColor = System.Drawing.Color.White;
			this.cmbGrupo.DisplayLayout.Appearance = appearance3;
			ultraGridColumn1.Header.VisiblePosition = 6;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 7;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 336;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 26;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 27;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 57;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 237;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
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
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.cmbGrupo.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbGrupo.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.cmbGrupo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbGrupo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupo.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			this.cmbGrupo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbGrupo.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupo.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.cmbGrupo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(120, 18);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(304, 21);
			this.cmbGrupo.TabIndex = 99;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(24, 18);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(39, 16);
			this.label2.TabIndex = 100;
			this.label2.Text = "Grupo:";
			// 
			// cvdArticuloGrupo
			// 
			this.cvdArticuloGrupo.BindingContextCtrl = this;
			this.cvdArticuloGrupo.DataSet = this.cdsArticuloTabla;
			this.cvdArticuloGrupo.Sort = "Grupo";
			this.cvdArticuloGrupo.TableName = "";
			this.cvdArticuloGrupo.TableViewName = "ArticuloGrupo";
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
			// 
			// cdsMarca
			// 
			this.cdsMarca.BindingContextCtrl = this;
			this.cdsMarca.DataLibrary = "LibFacturacion";
			this.cdsMarca.DataLibraryUrl = "";
			this.cdsMarca.DataSetDef = "dsArticuloTabla";
			this.cdsMarca.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMarca.Name = "cdsMarca";
			this.cdsMarca.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarca.SchemaDef = null;
			// 
			// txtIdMarca
			// 
			this.txtIdMarca.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdMarca.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarca, "ArticuloMarca.idMarca"));
			this.txtIdMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdMarca.Enabled = false;
			this.txtIdMarca.Location = new System.Drawing.Point(408, 104);
			this.txtIdMarca.Name = "txtIdMarca";
			this.txtIdMarca.PromptChar = ' ';
			this.txtIdMarca.Size = new System.Drawing.Size(24, 21);
			this.txtIdMarca.TabIndex = 235;
			this.txtIdMarca.Visible = false;
			// 
			// MarcasDeArticulo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(448, 134);
			this.Controls.Add(this.txtIdMarca);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.label2);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "MarcasDeArticulo";
			this.Text = "Marcas de artículos";
			this.Load += new System.EventHandler(this.MarcasDeArticulo_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cvdArticuloGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMarca)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drMarca = cdsMarca.TableViews["ArticuloMarca"].Rows[0];
			}
			catch{};
		}

		private void MarcasDeArticulo_Load(object sender, System.EventArgs e)
		{
			barraDato1.PosUltima();
			barraDato1.HabilitaControles(false);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
			this.cmbGrupo.Select();
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

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();
		}
	}
}
