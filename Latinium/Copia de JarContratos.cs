using System;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de JarContratos.
	/// </summary>
	public class JarContratos : DevExpress.XtraEditors.XtraForm
	{
//		private bool bNuevo = true;
		private C1.Data.C1DataSet cdsContratos;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbInteres;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private DevExpress.XtraEditors.SimpleButton btGrupo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor6;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor7;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCant1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaFinquito;
		private Infragistics.Win.Misc.UltraButton btPagos;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtId;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optSistema;
		private System.Windows.Forms.Label label2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public JarContratos()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idContratos = 0;
		public JarContratos(int idContrato)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
				idContratos = idContrato;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("JarContratosInteres", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJarContratosInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(JarContratos));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.cdsContratos = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cmbInteres = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.btGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.ultraNumericEditor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor7 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCant1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFechaFinquito = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btPagos = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.txtId = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optSistema = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label2 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCant1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFinquito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtId)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optSistema)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsContratos
			// 
			this.cdsContratos.BindingContextCtrl = this;
			this.cdsContratos.DataLibrary = "LibFacturacion";
			this.cdsContratos.DataLibraryUrl = "";
			this.cdsContratos.DataSetDef = "dsJarContratos";
			this.cdsContratos.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsContratos.Name = "cdsContratos";
			this.cdsContratos.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContratos.SchemaDef = null;
			this.cdsContratos.BeforeFill += new C1.Data.FillEventHandler(this.cdsContratos_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_JarContratos";
			this.barraDato1.DataNombreId = "IdJarContratos";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsContratos;
			this.barraDato1.DataTabla = "JarContratos";
			this.barraDato1.DataTablaHija = "JarContratosPagos";
			this.barraDato1.Location = new System.Drawing.Point(24, 264);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 127;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cmbInteres
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbInteres.Appearance = appearance1;
			this.cmbInteres.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbInteres.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbInteres.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.IdGrupoInteres"));
			this.cmbInteres.DataMember = "JarContratosInteres";
			this.cmbInteres.DataSource = this.cdsContratos;
			appearance2.BackColor = System.Drawing.Color.White;
			this.cmbInteres.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbInteres.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.cmbInteres.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbInteres.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.cmbInteres.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbInteres.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbInteres.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			this.cmbInteres.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbInteres.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbInteres.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.cmbInteres.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbInteres.DisplayMember = "Interes";
			this.cmbInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbInteres.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbInteres.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbInteres.Location = new System.Drawing.Point(336, 224);
			this.cmbInteres.Name = "cmbInteres";
			this.cmbInteres.Size = new System.Drawing.Size(112, 21);
			this.cmbInteres.TabIndex = 236;
			this.cmbInteres.ValueMember = "idJarContratosInteres";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.Location = new System.Drawing.Point(248, 224);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel6.TabIndex = 235;
			this.ultraLabel6.Text = "Interes Aplica";
			// 
			// btGrupo
			// 
			this.btGrupo.Image = ((System.Drawing.Image)(resources.GetObject("btGrupo.Image")));
			this.btGrupo.Location = new System.Drawing.Point(448, 224);
			this.btGrupo.Name = "btGrupo";
			this.btGrupo.Size = new System.Drawing.Size(24, 23);
			this.btGrupo.TabIndex = 234;
			this.btGrupo.ToolTip = "Grupo de artículos";
			this.btGrupo.Click += new System.EventHandler(this.btGrupo_Click);
			// 
			// ultraNumericEditor5
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor5.Appearance = appearance7;
			this.ultraNumericEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.SaldoInicial"));
			this.ultraNumericEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor5.FormatString = "n2";
			this.ultraNumericEditor5.Location = new System.Drawing.Point(336, 192);
			this.ultraNumericEditor5.Name = "ultraNumericEditor5";
			this.ultraNumericEditor5.NullText = "0";
			this.ultraNumericEditor5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor5.PromptChar = ' ';
			this.ultraNumericEditor5.ReadOnly = true;
			this.ultraNumericEditor5.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor5.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor5.TabIndex = 233;
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.Location = new System.Drawing.Point(248, 192);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel11.TabIndex = 232;
			this.ultraLabel11.Text = "Saldo";
			// 
			// ultraNumericEditor6
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor6.Appearance = appearance8;
			this.ultraNumericEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.NroCuotas"));
			this.ultraNumericEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor6.FormatString = "n2";
			this.ultraNumericEditor6.Location = new System.Drawing.Point(336, 160);
			this.ultraNumericEditor6.Name = "ultraNumericEditor6";
			this.ultraNumericEditor6.NullText = "0";
			this.ultraNumericEditor6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor6.PromptChar = ' ';
			this.ultraNumericEditor6.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor6.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor6.TabIndex = 231;
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.Location = new System.Drawing.Point(248, 160);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel12.TabIndex = 230;
			this.ultraLabel12.Text = "# de Cuotas";
			// 
			// ultraNumericEditor7
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor7.Appearance = appearance9;
			this.ultraNumericEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.Cuota"));
			this.ultraNumericEditor7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor7.FormatString = "n2";
			this.ultraNumericEditor7.Location = new System.Drawing.Point(336, 128);
			this.ultraNumericEditor7.Name = "ultraNumericEditor7";
			this.ultraNumericEditor7.NullText = "0";
			this.ultraNumericEditor7.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor7.PromptChar = ' ';
			this.ultraNumericEditor7.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor7.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor7.TabIndex = 229;
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.Location = new System.Drawing.Point(248, 128);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel13.TabIndex = 228;
			this.ultraLabel13.Text = "Cuota Inicial";
			// 
			// ultraNumericEditor3
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor3.Appearance = appearance10;
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.ValorCuota"));
			this.ultraNumericEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor3.FormatString = "n2";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(104, 224);
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.NullText = "0";
			this.ultraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(128, 21);
			this.ultraNumericEditor3.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor3.TabIndex = 227;
			this.ultraNumericEditor3.ValueChanged += new System.EventHandler(this.ultraNumericEditor3_ValueChanged);
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.Location = new System.Drawing.Point(24, 224);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel9.TabIndex = 226;
			this.ultraLabel9.Text = "Valor Cuota";
			// 
			// ultraNumericEditor2
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor2.Appearance = appearance11;
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.Descuento"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "n2";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(104, 192);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NullText = "0";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(128, 21);
			this.ultraNumericEditor2.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor2.TabIndex = 225;
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.Location = new System.Drawing.Point(24, 192);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel8.TabIndex = 224;
			this.ultraLabel8.Text = "Descuento";
			// 
			// ultraNumericEditor1
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance12;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.Interes"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "n2";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(104, 160);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NullText = "0";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 223;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.Location = new System.Drawing.Point(24, 160);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel7.TabIndex = 222;
			this.ultraLabel7.Text = "Interes";
			// 
			// txtCant1
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCant1.Appearance = appearance13;
			this.txtCant1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.Valor"));
			this.txtCant1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCant1.FormatString = "n2";
			this.txtCant1.Location = new System.Drawing.Point(104, 128);
			this.txtCant1.Name = "txtCant1";
			this.txtCant1.NullText = "0";
			this.txtCant1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCant1.PromptChar = ' ';
			this.txtCant1.Size = new System.Drawing.Size(128, 21);
			this.txtCant1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCant1.TabIndex = 221;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(248, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 219;
			this.label1.Text = "Pin";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.Location = new System.Drawing.Point(24, 128);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel5.TabIndex = 218;
			this.ultraLabel5.Text = "Valor Contrato";
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.Location = new System.Drawing.Point(24, 96);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel4.TabIndex = 217;
			this.ultraLabel4.Text = "Nombre";
			// 
			// ultraTextEditor3
			// 
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.Nombre"));
			this.ultraTextEditor3.Location = new System.Drawing.Point(104, 96);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(344, 21);
			this.ultraTextEditor3.TabIndex = 216;
			this.ultraTextEditor3.Tag = "Pin";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.Pin"));
			this.ultraTextEditor2.Location = new System.Drawing.Point(336, 64);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(112, 21);
			this.ultraTextEditor2.TabIndex = 214;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Location = new System.Drawing.Point(248, 32);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(88, 24);
			this.ultraLabel2.TabIndex = 212;
			this.ultraLabel2.Text = "Fecha Contrato";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Location = new System.Drawing.Point(24, 32);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(80, 24);
			this.ultraLabel1.TabIndex = 211;
			this.ultraLabel1.Text = "# de Solicitud";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.NumSolicitud"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(104, 32);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor1.TabIndex = 210;
			// 
			// cmbFechaFinquito
			// 
			this.cmbFechaFinquito.AllowMonthSelection = true;
			this.cmbFechaFinquito.AllowWeekSelection = true;
			this.cmbFechaFinquito.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaFinquito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.FechaContrato"));
			dateButton1.Caption = "Today";
			this.cmbFechaFinquito.DateButtons.Add(dateButton1);
			this.cmbFechaFinquito.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaFinquito.Format = "dd/MMM/yyyy";
			this.cmbFechaFinquito.Location = new System.Drawing.Point(336, 32);
			this.cmbFechaFinquito.Name = "cmbFechaFinquito";
			this.cmbFechaFinquito.NonAutoSizeHeight = 23;
			this.cmbFechaFinquito.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaFinquito.SpinButtonsVisible = true;
			this.cmbFechaFinquito.TabIndex = 213;
			this.cmbFechaFinquito.Value = new System.DateTime(2010, 5, 28, 0, 0, 0, 0);
			// 
			// btPagos
			// 
			this.btPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPagos.Location = new System.Drawing.Point(384, 264);
			this.btPagos.Name = "btPagos";
			this.btPagos.Size = new System.Drawing.Size(32, 24);
			this.btPagos.TabIndex = 237;
			this.btPagos.Text = "P";
			this.btPagos.Click += new System.EventHandler(this.btPagos_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.Location = new System.Drawing.Point(344, 264);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(32, 24);
			this.ultraButton1.TabIndex = 238;
			this.ultraButton1.Text = "S";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton2.Location = new System.Drawing.Point(424, 264);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(32, 24);
			this.ultraButton2.TabIndex = 239;
			this.ultraButton2.Text = "F";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// txtId
			// 
			this.txtId.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "JarContratos.IdJarContratos"));
			this.txtId.Location = new System.Drawing.Point(424, 0);
			this.txtId.Name = "txtId";
			this.txtId.Size = new System.Drawing.Size(0, 21);
			this.txtId.TabIndex = 240;
			// 
			// optSistema
			// 
			this.optSistema.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContratos, "_JarContratos.Status"));
			this.optSistema.ItemAppearance = appearance14;
			valueListItem1.DataValue = "AC";
			valueListItem1.DisplayText = "AC";
			valueListItem2.DataValue = "PC";
			valueListItem2.DisplayText = "PC";
			this.optSistema.Items.Add(valueListItem1);
			this.optSistema.Items.Add(valueListItem2);
			this.optSistema.ItemSpacingVertical = 6;
			this.optSistema.Location = new System.Drawing.Point(104, 64);
			this.optSistema.Name = "optSistema";
			this.optSistema.Size = new System.Drawing.Size(128, 24);
			this.optSistema.TabIndex = 241;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(24, 64);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 16);
			this.label2.TabIndex = 242;
			this.label2.Text = "Status";
			// 
			// JarContratos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 301);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.optSistema);
			this.Controls.Add(this.txtId);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btPagos);
			this.Controls.Add(this.cmbInteres);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.btGrupo);
			this.Controls.Add(this.ultraNumericEditor5);
			this.Controls.Add(this.ultraLabel11);
			this.Controls.Add(this.ultraNumericEditor6);
			this.Controls.Add(this.ultraLabel12);
			this.Controls.Add(this.ultraNumericEditor7);
			this.Controls.Add(this.ultraLabel13);
			this.Controls.Add(this.ultraNumericEditor3);
			this.Controls.Add(this.ultraLabel9);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.txtCant1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.ultraTextEditor3);
			this.Controls.Add(this.ultraTextEditor2);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.cmbFechaFinquito);
			this.Controls.Add(this.barraDato1);
			this.Name = "JarContratos";
			this.Text = "JarContratos";
			this.Load += new System.EventHandler(this.JarContratos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsContratos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCant1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFinquito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtId)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optSistema)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraTextEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraLabel2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsContratos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsContratos.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void JarContratos_Load(object sender, System.EventArgs e)
		{
			if (idContratos > 0)
				Busqueda(idContratos);
			else
				barraDato1.PosUltima();
			barraDato1.HabilitaControles(false);
			
		}
		private void Busqueda(int idBusca)
		{
			barraDato1.IdRegistro = idBusca;
			barraDato1.ProximoId(5);
		}
		private void HabilitaTab()
		{
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			
			string stSelect = "Select Count(*) From JarContratosPagos "
				+ "Where idJarContratos = " + txtId.Value.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsContratos, stSelect, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede borrar Contratos porque existen "
					+ iCuenta.ToString() + " registros en pagos", "Información");
				return;
			}
			if (MessageBox.Show("¿Desea Borrar datos de éste contrato?",
				"Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
				== DialogResult.No) return;
			stSelect = "Delete JarContratos  "
				+ "Where idJarContratos = " + txtId.Value.ToString();
			Funcion.sEscalarSQL(cdsContratos, stSelect, true);

			barraDato1.ProximoId(2);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			optSistema.Value = "AC";
			cmbFechaFinquito.Value = DateTime.Today;
			barraDato1.CrearRegistro();

		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();

		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{

		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			JarContratosRep miReporte = new JarContratosRep();
			miReporte.MdiParent = MdiParent;
			miReporte.Show();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
//			JarContratosBuscar miJarContratos = new JarContratosBuscar();
//			miJarContratos.MdiParent = MdiParent;
//			miJarContratos.Show();

			using (JarContratosBuscar miBusqueda = new JarContratosBuscar())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					try
					{
						int id = int.Parse(miBusqueda.idBusca.Text);
						Busqueda(id);
					}
					catch{}
				}
			}
	
		}

		private void ultraTextEditor3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btGrupo_Click(object sender, System.EventArgs e)
		{
			using(JarContratosInteres miInteres = new JarContratosInteres())
			{
				miInteres.ShowDialog();
			}
			this.cmbInteres.Refresh();

		}

		private void btPagos_Click(object sender, System.EventArgs e)
		{
//			using (JarContratosBuscar miBusqueda = new JarContratosPagos(idContratos))
//			{
//				if (miBusqueda.ShowDialog() == DialogResult.OK)
//				{
//					try
//					{
//						int id = int.Parse(miBusqueda.idBusca.Text);
//						Busqueda(id);
//					}
//					catch{}
//				}
//			}
			idContratos = int.Parse(txtId.Value.ToString());
			JarContratosPagos miContratopag = new JarContratosPagos(idContratos);
			miContratopag.MdiParent = MdiParent;
			miContratopag.Show();
			barraDato1.Refresh();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			JarContratosSubida miContratosub = new JarContratosSubida();
			miContratosub.MdiParent = MdiParent;
			miContratosub.Show();
			this.cdsContratos.Clear();
			this.cdsContratos.Fill();
			this.cmbInteres.Refresh();
			barraDato1.PosUltima();
			Cursor = Cursors.Default;

		}

		private void ultraNumericEditor3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			idContratos = int.Parse(txtId.Value.ToString());
			if (DialogResult.No == MessageBox.Show("Desea realizar una factura", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;
			string stExec = "Exec  JarContratosFactura " + idContratos.ToString();
			Funcion.EjecutaSQL(cdsContratos,stExec); 
			MessageBox.Show("Proceso Realizada", "Información");
			if (idContratos > 0)
				Busqueda(idContratos);
			else
				barraDato1.PosUltima();

		}


	}
}
