using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Form2.
	/// </summary>
	public class Form2 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label11;
		private C1.Data.C1DataSet cdsEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label Observacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdCasa;
		private C1.Data.C1DataSet cdsObservacion;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbColorcombo;
		private DevExpress.XtraEditors.LookUpEdit cmbCliente;
		private System.Windows.Forms.Label lblCliente;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form2()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PointEstado", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PointObservaciones", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PointObservaciones", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado", -1, "cmbColorcombo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("PointEstado", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			this.label11 = new System.Windows.Forms.Label();
			this.cdsEstado = new C1.Data.C1DataSet();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.Observacion = new System.Windows.Forms.Label();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsObservacion = new C1.Data.C1DataSet();
			this.grdCasa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbColorcombo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCliente = new DevExpress.XtraEditors.LookUpEdit();
			this.lblCliente = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsObservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColorcombo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(24, 8);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(39, 16);
			this.label11.TabIndex = 50;
			this.label11.Text = "Estado";
			// 
			// cdsEstado
			// 
			this.cdsEstado.BindingContextCtrl = this;
			this.cdsEstado.DataLibrary = "LibFacturacion";
			this.cdsEstado.DataLibraryUrl = "";
			this.cdsEstado.DataSetDef = "dsEstado";
			this.cdsEstado.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEstado.Name = "cdsEstado";
			this.cdsEstado.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsEstado.SchemaDef = null;
			this.cdsEstado.BeforeFill += new C1.Data.FillEventHandler(this.cdsEstado_BeforeFill);
			// 
			// cmbEstado
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance1;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataMember = "PointEstado";
			this.cmbEstado.DataSource = this.cdsEstado;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(72, 8);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(128, 21);
			this.cmbEstado.TabIndex = 52;
			this.cmbEstado.ValueMember = "idEstado";
			// 
			// Observacion
			// 
			this.Observacion.AutoSize = true;
			this.Observacion.Location = new System.Drawing.Point(24, 40);
			this.Observacion.Name = "Observacion";
			this.Observacion.Size = new System.Drawing.Size(62, 16);
			this.Observacion.TabIndex = 53;
			this.Observacion.Text = "Observcion";
			// 
			// ultraCombo2
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo2.Appearance = appearance2;
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DataMember = "PointObservaciones";
			this.ultraCombo2.DataSource = this.cdsEstado;
			ultraGridBand2.AddButtonCaption = "PointEstado";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.ultraCombo2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraCombo2.DisplayMember = "Observaciones";
			this.ultraCombo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(88, 40);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo2.TabIndex = 55;
			this.ultraCombo2.ValueMember = "idObservaciones";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Pointobservaciones";
			this.barraDato1.DataNombreId = "idObservaciones";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsObservacion;
			this.barraDato1.DataTabla = "PointObservacones";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 264);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 56;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			// 
			// cdsObservacion
			// 
			this.cdsObservacion.BindingContextCtrl = this;
			this.cdsObservacion.DataLibrary = "LibFacturacion";
			this.cdsObservacion.DataLibraryUrl = "";
			this.cdsObservacion.DataSetDef = "dsEstado";
			this.cdsObservacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsObservacion.Name = "cdsObservacion";
			this.cdsObservacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsObservacion.SchemaDef = null;
			// 
			// grdCasa
			// 
			this.grdCasa.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdCasa.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdCasa.DataMember = "PointObservaciones";
			this.grdCasa.DataSource = this.cdsObservacion;
			appearance3.BackColor = System.Drawing.Color.White;
			this.grdCasa.DisplayLayout.Appearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
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
			this.grdCasa.Location = new System.Drawing.Point(40, 96);
			this.grdCasa.Name = "grdCasa";
			this.grdCasa.Size = new System.Drawing.Size(384, 136);
			this.grdCasa.TabIndex = 122;
			// 
			// cmbColorcombo
			// 
			this.cmbColorcombo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColorcombo.DataMember = "PointEstado";
			this.cmbColorcombo.DataSource = this.cdsEstado;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbColorcombo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbColorcombo.DisplayMember = "Estado";
			this.cmbColorcombo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColorcombo.Location = new System.Drawing.Point(296, 144);
			this.cmbColorcombo.Name = "cmbColorcombo";
			this.cmbColorcombo.Size = new System.Drawing.Size(200, 96);
			this.cmbColorcombo.TabIndex = 123;
			this.cmbColorcombo.Text = "ultraDropDown1";
			this.cmbColorcombo.ValueMember = "idEstado";
			this.cmbColorcombo.Visible = false;
			// 
			// cmbCliente
			// 
			this.cmbCliente.Location = new System.Drawing.Point(152, 72);
			this.cmbCliente.Name = "cmbCliente";
			// 
			// cmbCliente.Properties
			// 
			this.cmbCliente.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.Ascending),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("idCliente", "", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbCliente.Properties.DisplayMember = "Nombre";
			this.cmbCliente.Properties.MaxLength = 400;
			this.cmbCliente.Properties.NullText = "";
			this.cmbCliente.Properties.PopupWidth = 400;
			this.cmbCliente.Properties.ValueMember = "idCliente";
			this.cmbCliente.Size = new System.Drawing.Size(248, 20);
			this.cmbCliente.TabIndex = 143;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(104, 72);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(43, 16);
			this.lblCliente.TabIndex = 142;
			this.lblCliente.Text = "Cliente:";
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
			// 
			// Form2
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 294);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.cmbColorcombo);
			this.Controls.Add(this.grdCasa);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraCombo2);
			this.Controls.Add(this.Observacion);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.label11);
			this.Name = "Form2";
			this.Text = "Form2";
			this.Load += new System.EventHandler(this.Form2_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsObservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColorcombo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

     C1.Data.C1DataRow drEstado;
		private void cdsEstado_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsEstado.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Form2_Load(object sender, System.EventArgs e)
		{
			
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drEstado = cdsEstado.TableViews["Observaciones"].Rows[0]; //fila con la trabajo
			}
			catch{};
			
		}
		private void ClienteCarga()
		{
//			if (drSeteoF["VerNombreLista"] != DBNull.Value && (bool) drSeteoF["VerNombreLista"])
//			{
				string stExec = "Exec ClienteCargaNombre " + 4;
				cmbCliente.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
				//txtNombre.Width = 0;
//			}
//			else
//				cmbCliente.Width = 0;
//
//			cmbRecibe.DataSource = 	Funcion.dvProcedimiento(cdsCompra,
//				"Select Nombre, Codigo, idCliente From Cliente Where Entrega = 1 and (Proveedor = "+ProvRecibFabara+" or Ambos = 1)");
		}
	}
}
