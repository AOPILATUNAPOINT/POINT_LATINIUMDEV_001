using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for GruposSubgrupos.
	/// </summary>
	public class GruposSubgrupos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsGrupoSubgrupo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataSet cdsArticuloTabla;
		private C1.Data.C1DataView cdvSubGrupo;
		private C1.Data.C1DataView cvdArticuloGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubgrupo;

		C1.Data.C1DataRow drGrupoSubGrupo;

		public GruposSubgrupos()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSubGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsGrupoSubgrupo = new C1.Data.C1DataSet();
			this.cvdArticuloGrupo = new C1.Data.C1DataView();
			this.cdsArticuloTabla = new C1.Data.C1DataSet();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbSubgrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSubGrupo = new C1.Data.C1DataView();
			this.label1 = new System.Windows.Forms.Label();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoSubgrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cvdArticuloGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubgrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbGrupo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbGrupo.Appearance = appearance1;
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoSubgrupo, "GrupoSubgrupo.idGrupoArticulo"));
			this.cmbGrupo.DataSource = this.cvdArticuloGrupo;
			appearance2.BackColor = System.Drawing.Color.White;
			this.cmbGrupo.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 305;
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
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
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
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.cmbGrupo.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbGrupo.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.cmbGrupo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbGrupo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupo.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			this.cmbGrupo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbGrupo.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbGrupo.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.cmbGrupo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(101, 14);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(315, 21);
			this.cmbGrupo.TabIndex = 220;
			this.cmbGrupo.ValueMember = "idGrupoArticulo";
			this.cmbGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbGrupo_InitializeLayout);
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
			this.cdsArticuloTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticuloTabla_BeforeFill);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(34, 14);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(39, 16);
			this.label9.TabIndex = 221;
			this.label9.Text = "Grupo:";
			// 
			// cmbSubgrupo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubgrupo.Appearance = appearance7;
			this.cmbSubgrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubgrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubgrupo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoSubgrupo, "GrupoSubgrupo.idSubgrupo"));
			this.cmbSubgrupo.DataSource = this.cdvSubGrupo;
			appearance8.BackColor = System.Drawing.Color.White;
			this.cmbSubgrupo.DisplayLayout.Appearance = appearance8;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 315;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbSubgrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.cmbSubgrupo.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbSubgrupo.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.cmbSubgrupo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbSubgrupo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbSubgrupo.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			this.cmbSubgrupo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSubgrupo.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbSubgrupo.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.cmbSubgrupo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbSubgrupo.DisplayMember = "SubGrupo";
			this.cmbSubgrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubgrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubgrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubgrupo.Location = new System.Drawing.Point(101, 46);
			this.cmbSubgrupo.Name = "cmbSubgrupo";
			this.cmbSubgrupo.Size = new System.Drawing.Size(315, 21);
			this.cmbSubgrupo.TabIndex = 222;
			this.cmbSubgrupo.ValueMember = "idSubGrupo";
			// 
			// cdvSubGrupo
			// 
			this.cdvSubGrupo.BindingContextCtrl = this;
			this.cdvSubGrupo.DataSet = this.cdsArticuloTabla;
			this.cdvSubGrupo.Sort = "SubGrupo";
			this.cdvSubGrupo.TableName = "";
			this.cdvSubGrupo.TableViewName = "ArticuloSubGrupo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(34, 46);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 223;
			this.label1.Text = "Subgrupo:";
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoSubgrupo, "GrupoSubgrupo.idGrupoSubGrupo"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Location = new System.Drawing.Point(386, 78);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(24, 22);
			this.txtIdCompra.TabIndex = 224;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "GrupoSubGrupo";
			this.barraDato1.DataNombreId = "idGrupoSubgrupo";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsGrupoSubgrupo;
			this.barraDato1.DataTabla = "GrupoSubGrupo";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(80, 78);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 225;
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
			// GruposSubgrupos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(450, 118);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.cmbSubgrupo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.cmbGrupo);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "GruposSubgrupos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Grupos - Subgrupos";
			this.Load += new System.EventHandler(this.GruposSubgrupos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoSubgrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cvdArticuloGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticuloTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubgrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void GruposSubgrupos_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.PosUltima();
			barraDato1.HabilitaControles(false);
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drGrupoSubGrupo = cdsGrupoSubgrupo.TableViews["GrupoSubGrupo"].Rows[0];
			}
			catch{};
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
			if (drGrupoSubGrupo["idGrupoArticulo"] == DBNull.Value)
			{
				MessageBox.Show("Seleccione un grupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbGrupo.Select();
				return;
			}
			
			if (drGrupoSubGrupo["idSubGrupo"] == DBNull.Value)
			{
				MessageBox.Show("Seleccione un Subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cmbSubgrupo.Select();
				return;
			}

			string sSQL = string.Format("Select Count(*) From GrupoSubGrupo Where idGrupoArticulo = {0} And  idSubGrupo = {1}", (int)drGrupoSubGrupo["idGrupoArticulo"], (int)drGrupoSubGrupo["idSubGrupo"]);
			int iCuenta = 0;
			iCuenta = Funcion.iEscalarSQL(cdsArticuloTabla, sSQL);
			if (iCuenta > 0)
			{
				MessageBox.Show("Registro existente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			this.barraDato1.GrabaRegistro();
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

		private void cmbGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();
		}
	}
}

