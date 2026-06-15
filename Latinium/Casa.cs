using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Casa.
	/// </summary>
	public class Casa : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label22;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private System.Windows.Forms.Label label30;
		private C1.Data.C1DataSet cdsCasa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCasa;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbColor;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoPrecio;
		private C1.Data.C1DataSet cdsCasa2;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdCasa;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Casa()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloColor", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloColor", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Casa", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCasa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCasa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor", -1, "cmbTipoPrecio");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.label22 = new System.Windows.Forms.Label();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.txtCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsCasa = new C1.Data.C1DataSet();
			this.label30 = new System.Windows.Forms.Label();
			this.cmbColor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsCasa2 = new C1.Data.C1DataSet();
			this.cmbTipoPrecio = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.grdCasa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCasa2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCasa)).BeginInit();
			this.SuspendLayout();
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(88, 24);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(43, 16);
			this.label22.TabIndex = 116;
			this.label22.Text = "Mi Dato";
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
			// txtCasa
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCasa.Appearance = appearance1;
			this.txtCasa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCasa, "Casa.NombreCasa"));
			this.txtCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCasa.Location = new System.Drawing.Point(144, 24);
			this.txtCasa.Name = "txtCasa";
			this.txtCasa.Size = new System.Drawing.Size(152, 21);
			this.txtCasa.TabIndex = 115;
			// 
			// cdsCasa
			// 
			this.cdsCasa.BindingContextCtrl = this;
			this.cdsCasa.DataLibrary = "LibFacturacion";
			this.cdsCasa.DataLibraryUrl = "";
			this.cdsCasa.DataSetDef = "dsCasa";
			this.cdsCasa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCasa.Name = "cdsCasa";
			this.cdsCasa.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCasa.SchemaDef = null;
			this.cdsCasa.BeforeFill += new C1.Data.FillEventHandler(this.cdsCasa_BeforeFill);
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(88, 56);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(33, 16);
			this.label30.TabIndex = 114;
			this.label30.Text = "Clase";
			// 
			// cmbColor
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbColor.Appearance = appearance2;
			this.cmbColor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbColor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCasa, "Casa.idColor"));
			this.cmbColor.DataMember = "ArticuloColor";
			this.cmbColor.DataSource = this.cdsArticuloTabla;
			this.cmbColor.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 58;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 99;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbColor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbColor.DisplayMember = "Color";
			this.cmbColor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbColor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColor.Location = new System.Drawing.Point(144, 56);
			this.cmbColor.Name = "cmbColor";
			this.cmbColor.Size = new System.Drawing.Size(176, 21);
			this.cmbColor.TabIndex = 121;
			this.cmbColor.ValueMember = "idColor";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Casa";
			this.barraDato1.DataNombreId = "idCasa";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCasa;
			this.barraDato1.DataTabla = "Casa";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(232, 352);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 122;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsCasa2
			// 
			this.cdsCasa2.BindingContextCtrl = this;
			this.cdsCasa2.DataLibrary = "LibFacturacion";
			this.cdsCasa2.DataLibraryUrl = "";
			this.cdsCasa2.DataSetDef = "dsCasa";
			this.cdsCasa2.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCasa2.Name = "cdsCasa2";
			this.cdsCasa2.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCasa2.SchemaDef = null;
			this.cdsCasa2.BeforeFill += new C1.Data.FillEventHandler(this.cdsCasa2_BeforeFill);
			// 
			// cmbTipoPrecio
			// 
			this.cmbTipoPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoPrecio.DataMember = "ArticuloColor";
			this.cmbTipoPrecio.DataSource = this.cdsArticuloTabla;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbTipoPrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoPrecio.DisplayMember = "Color";
			this.cmbTipoPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoPrecio.Location = new System.Drawing.Point(384, 224);
			this.cmbTipoPrecio.Name = "cmbTipoPrecio";
			this.cmbTipoPrecio.Size = new System.Drawing.Size(144, 112);
			this.cmbTipoPrecio.TabIndex = 125;
			this.cmbTipoPrecio.Text = "ultraDropDown1";
			this.cmbTipoPrecio.ValueMember = "idColor";
			this.cmbTipoPrecio.Visible = false;
			// 
			// grdCasa
			// 
			this.grdCasa.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCasa.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCasa.DataMember = "Casa";
			this.grdCasa.DataSource = this.cdsCasa2;
			appearance3.BackColor = System.Drawing.Color.White;
			this.grdCasa.DisplayLayout.Appearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand3.SummaryFooterCaption = "";
			this.grdCasa.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.grdCasa.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdCasa.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdCasa.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.grdCasa.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdCasa.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCasa.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCasa.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdCasa.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.grdCasa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdCasa.Location = new System.Drawing.Point(16, 120);
			this.grdCasa.Name = "grdCasa";
			this.grdCasa.Size = new System.Drawing.Size(608, 152);
			this.grdCasa.TabIndex = 126;
			// 
			// Casa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(640, 382);
			this.Controls.Add(this.grdCasa);
			this.Controls.Add(this.cmbTipoPrecio);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbColor);
			this.Controls.Add(this.label30);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.txtCasa);
			this.Name = "Casa";
			this.Text = "Casa";
			this.Load += new System.EventHandler(this.Casa_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCasa2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCasa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drCasa;

		private void cdsCasa_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCasa.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();
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

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drCasa = cdsCasa.TableViews["Casa"].Rows[0];
			}
			catch{};
			
		}

		private void Casa_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.HabilitaControles(false);
			this.barraDato1.PosUltima();
		}

		private void cdsCasa2_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCasa2.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}
	}
}
