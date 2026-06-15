using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de PresupuestoCompras.
	/// </summary>
	public class PresupuestoCompras : DevExpress.XtraEditors.XtraForm
	{
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsPresupuestoC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public PresupuestoCompras()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_P_PresupuestoCapitulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCapitulo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_PresupCapit - P_PresupSubCap");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("P_PresupCapit - P_PresupSubCap", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capitulo-Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Capitulo-Detalle", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoCDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rubro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PUnitario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubCapitulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Maquinaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManoObra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Materiales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Indirectos");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsPresupuestoC = new C1.Data.C1DataSet();
			this.txtObra = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_P_Presupuesto";
			this.barraDato1.DataNombreId = "idPresupuestoC";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsPresupuestoC;
			this.barraDato1.DataTabla = "P_Presupuesto";
			this.barraDato1.DataTablaHija = "P_PresupuestoCapitulo";
			this.barraDato1.Location = new System.Drawing.Point(10, 433);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(364, 30);
			this.barraDato1.TabIndex = 11;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsPresupuestoC
			// 
			this.cdsPresupuestoC.BindingContextCtrl = this;
			this.cdsPresupuestoC.DataLibrary = "LibFacturacion";
			this.cdsPresupuestoC.DataLibraryUrl = "";
			this.cdsPresupuestoC.DataSetDef = "dsPresupuestoC";
			this.cdsPresupuestoC.EnforceConstraints = false;
			this.cdsPresupuestoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPresupuestoC.Name = "cdsPresupuestoC";
			this.cdsPresupuestoC.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPresupuestoC.SchemaDef = null;
			this.cdsPresupuestoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsPresupuestoC_BeforeFill);
			// 
			// txtObra
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObra.Appearance = appearance1;
			this.txtObra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPresupuestoC, "_P_Presupuesto.Obra"));
			this.txtObra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObra.Location = new System.Drawing.Point(86, 20);
			this.txtObra.Name = "txtObra";
			this.txtObra.Size = new System.Drawing.Size(260, 25);
			this.txtObra.TabIndex = 12;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(19, 20);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(58, 28);
			this.label1.TabIndex = 13;
			this.label1.Text = "OBRA:";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_P_Presupuesto.P_Presupuesto_Capitulo";
			this.ultraGrid1.DataSource = this.cdsPresupuestoC;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 293;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 852;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 175;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 526;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Width = 529;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 134;
			ultraGridColumn10.Header.Caption = "Cuenta";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn10.Width = 71;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 105;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 163;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Width = 132;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Width = 111;
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 107;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Width = 115;
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Width = 133;
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 84;
			ultraGridColumn20.Header.VisiblePosition = 6;
			ultraGridColumn20.Width = 83;
			ultraGridColumn21.Header.VisiblePosition = 7;
			ultraGridColumn21.Width = 79;
			ultraGridColumn22.Header.VisiblePosition = 8;
			ultraGridColumn22.Width = 81;
			ultraGridColumn23.Header.VisiblePosition = 9;
			ultraGridColumn23.Width = 103;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(10, 59);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1185, 355);
			this.ultraGrid1.TabIndex = 14;
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// txtNotas
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance8;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPresupuestoC, "_P_Presupuesto.NotasObra"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(538, 20);
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(259, 25);
			this.txtNotas.TabIndex = 15;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(461, 20);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(57, 28);
			this.label2.TabIndex = 16;
			this.label2.Text = "NOTA:";
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.cdvCuenta;
			this.cmbCuenta.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Width = 154;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 4;
			ultraGridColumn28.Width = 199;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCuenta.DisplayMember = "Codigo";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(384, 216);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(355, 148);
			this.cmbCuenta.TabIndex = 129;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo = 0";
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.CaseSensitive = false;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// PresupuestoCompras
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(1204, 468);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtObra);
			this.Controls.Add(this.barraDato1);
			this.Name = "PresupuestoCompras";
			this.Text = "PresupuestoCompras";
			this.Load += new System.EventHandler(this.PresupuestoCompras_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsPresupuestoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cdsPresupuestoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		
			Funcion miFuncion = new Funcion();
			cdsPresupuestoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			
			barraDato1.CrearRegistro();
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

		private void PresupuestoCompras_Load(object sender, System.EventArgs e)
		{
			
			barraDato1.ProximoId(4);

		}
		private C1.Data.C1DataRow drPresupuesto;
		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drPresupuesto = cdsPresupuestoC.TableViews["P_Presupuesto"].Rows[0];
			}
			catch
			{}
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
	
			if (drPresupuesto == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string stSelect = "Select Count(*) from detcompra where idPresupuestoC = "+drPresupuesto["idPresupuestoC"];
			int ids = Funcion.iEscalarSQL(cdsPresupuestoC,stSelect);
			if(ids > 0 )
			{
				MessageBox.Show("Tiene registros Relacionados no se puee eliminar....");
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Borrar Presupuesto...", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

		barraDato1.BorraRegistro();
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
			{						
				if (e.Cell.Column.ToString()== "Maquinaria" || e.Cell.Column.ToString()== "ManoObra" 
					||e.Cell.Column.ToString()== "Materiales" || e.Cell.Column.ToString()== "Indirectos"  )	
				{
					double pU = (double) e.Cell.Row.Cells["PUnitario"].Value;
					double maq = (double) e.Cell.Row.Cells["Maquinaria"].Value;
					double manO = (double) e.Cell.Row.Cells["ManoObra"].Value;
					double mat = (double) e.Cell.Row.Cells["Materiales"].Value;
					double Ind = (double) e.Cell.Row.Cells["Indirectos"].Value;
					pU =  maq+ manO+ mat+ Ind;					
					e.Cell.Row.Cells["PUnitario"].Value = pU;

					double cant = (double) e.Cell.Row.Cells["Cantidad"].Value;
					double precio = (double) e.Cell.Row.Cells["PUnitario"].Value;

					e.Cell.Row.Cells["PTotal"].Value =  cant * precio;

				}
//				if (e.Cell.Column.ToString()== "Cantidad" || e.Cell.Column.ToString()== "PUnitario" )					
//				{
//					if(e.Cell.Row.Cells["Cantidad"].Value != DBNull.Value || e.Cell.Row.Cells["PUnitario"].Value != DBNull.Value)
//					{
//						double cant = (double) e.Cell.Row.Cells["Cantidad"].Value;
//						double precio = (double) e.Cell.Row.Cells["PUnitario"].Value;
//
//						e.Cell.Row.Cells["PTotal"].Value =  cant * precio;
//					}
//				}
			}
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);			
		}
	}
}
