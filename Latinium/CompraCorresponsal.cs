using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraComision.
	/// </summary>
	public class CompraCorresponsal : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsComision;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinGrid.UltraGrid grAsiento;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraCorresponsal()
		{
			InitializeComponent();
		}

		int IdCompra = 0;
		public CompraCorresponsal(int idCompra)
		{
			InitializeComponent();
			IdCompra=idCompra;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraCorresponsal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCorresponsal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", -1, true, "CompraCorresponsal", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			this.grAsiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsComision = new C1.Data.C1DataSet();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCalendarCombo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			this.SuspendLayout();
			// 
			// grAsiento
			// 
			this.grAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grAsiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.grAsiento.DataMember = "CompraCorresponsal";
			this.grAsiento.DataSource = this.cdsComision;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grAsiento.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "CompraComision";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.Caption = "Corresponsal";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn3.Width = 341;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 162;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.grAsiento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grAsiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grAsiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grAsiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.grAsiento.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grAsiento.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grAsiento.Location = new System.Drawing.Point(16, 40);
			this.grAsiento.Name = "grAsiento";
			this.grAsiento.Size = new System.Drawing.Size(552, 160);
			this.grAsiento.TabIndex = 0;
			this.grAsiento.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grAsiento_InitializeRow);
			this.grAsiento.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.grAsiento.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.grAsiento_BeforeRowInsert);
			this.grAsiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grAsiento_KeyDown);
			// 
			// cdsComision
			// 
			this.cdsComision.BindingContextCtrl = this;
			this.cdsComision.DataLibrary = "LibFacturacion";
			this.cdsComision.DataLibraryUrl = "";
			this.cdsComision.DataSetDef = "dsComision";
			this.cdsComision.EnforceConstraints = false;
			this.cdsComision.FillOnRequest = false;
			this.cdsComision.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsComision.Name = "cdsComision";
			this.cdsComision.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsComision.SchemaDef = null;
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = false;
			this.barraDatoSQL1.DataMember = "Compra";
			this.barraDatoSQL1.DataNombreId = "idCompra";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsCompra;
			this.barraDatoSQL1.DataTabla = "Compra";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(16, 216);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(96, 24);
			this.barraDatoSQL1.TabIndex = 1;
			this.barraDatoSQL1.VisibleBorrar = false;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Compra.idCompra"));
			this.txtIdCompra.Location = new System.Drawing.Point(352, 8);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(56, 21);
			this.txtIdCompra.TabIndex = 2;
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Compra.Numero"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(72, 8);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.ReadOnly = true;
			this.ultraTextEditor1.Size = new System.Drawing.Size(88, 21);
			this.ultraTextEditor1.TabIndex = 3;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(44, 14);
			this.ultraLabel1.TabIndex = 4;
			this.ultraLabel1.Text = "Numero";
			// 
			// ultraCalendarCombo1
			// 
			this.ultraCalendarCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Compra.Fecha"));
			dateButton1.Caption = "Today";
			this.ultraCalendarCombo1.DateButtons.Add(dateButton1);
			this.ultraCalendarCombo1.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo1.Location = new System.Drawing.Point(240, 8);
			this.ultraCalendarCombo1.Name = "ultraCalendarCombo1";
			this.ultraCalendarCombo1.NonAutoSizeHeight = 23;
			this.ultraCalendarCombo1.ReadOnly = true;
			this.ultraCalendarCombo1.Size = new System.Drawing.Size(104, 21);
			this.ultraCalendarCombo1.TabIndex = 5;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(192, 8);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(36, 14);
			this.ultraLabel2.TabIndex = 6;
			this.ultraLabel2.Text = "Fecha";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.EnforceConstraints = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "Entrega = 1";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.cdvCliente;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 212;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraDropDown1.DisplayMember = "Nombre";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(216, 88);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(344, 80);
			this.ultraDropDown1.TabIndex = 7;
			this.ultraDropDown1.ValueMember = "idCliente";
			this.ultraDropDown1.Visible = false;
			// 
			// CompraCorresponsal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(592, 254);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraCalendarCombo1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.grAsiento);
			this.Name = "CompraCorresponsal";
			this.Text = "Corresponsal de factura";
			this.Load += new System.EventHandler(this.CompraComision_Load);
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			MessageBox.Show("No se puede crear Corresponsales", "Información");
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.EditarRegistro();
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			cdsComision.Update();
			barraDatoSQL1.GrabaRegistro();
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			CompraCorrImprime miImprime = new CompraCorrImprime();
			miImprime.MdiParent=this.MdiParent;
			miImprime.Show();
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(cdsComision, "CompraCorresponsal", 
					" [idCompra] = " + txtIdCompra.Value.ToString());
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				cdsComision.Fill(fcTotal, false);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "error refresca");
			}
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		}

		private void CompraComision_Load(object sender, System.EventArgs e)
		{
			txtIdCompra.Width=0;
			if (IdCompra>0)
			{
				barraDatoSQL1.IdRegistro=IdCompra;
				barraDatoSQL1.ProximoId(5);
			}
			else
				barraDatoSQL1.ProximoId(4);

			barraDatoSQL1.HabilitaControles(false);
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCompra"].Value = (int) txtIdCompra.Value;
			e.Row.Cells["Valor"].Value=0;
		}

		private void grAsiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.Alt) 
				return;
			if (grAsiento.ActiveCell == null) return;
			switch(e.KeyValue)
			{
				case (int) Keys.PageDown:
					if(grAsiento.ActiveCell.DroppedDown == false)
						grAsiento.ActiveCell.DroppedDown = true;
					break;

				case (int) Keys.Up:
					if (grAsiento.ActiveCell.DroppedDown) return;

					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (grAsiento.ActiveCell.DroppedDown) return;
			
					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Enter:
					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
			}		
		}

		private void grAsiento_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
//			if (e.Row.Cells["Valor"].Value != DBNull.Value && e.Row.Cells["Porcentaje"].Value != DBNull.Value)
//			e.Row.Cells["Total"].Value = (double) e.Row.Cells["Valor"].Value * (double) e.Row.Cells["Porcentaje"].Value / 100.0;
		}

		private void cmbTotal_Click(object sender, System.EventArgs e)
		{
//			double dTotal = 0;
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
//			{
//				if (dr.Cells["Total"].Value != System.DBNull.Value)
//					dTotal += (double) dr.Cells["Total"].Value;
//			}
//			txtTotal.Value = dTotal;
		}

		private void grAsiento_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
		}

	}
}
