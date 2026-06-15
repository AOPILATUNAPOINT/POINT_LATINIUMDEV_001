using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ProductoMes.
	/// </summary>
	public class ProductoMes : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsProdSemana;
		private DevExpress.XtraEditors.DateEdit dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataView cdvPronostico;
		private C1.Data.C1DataSet cdsPronostico;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtMes;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ProductoMes()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ProdSemana", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProdSemana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PesoMes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaC1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaC2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaC3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaC4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaC5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaP1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaP2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaP3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaP4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SemanaP5");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cdsProdSemana = new C1.Data.C1DataSet();
			this.dtFecha = new DevExpress.XtraEditors.DateEdit();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cdvPronostico = new C1.Data.C1DataView();
			this.cdsPronostico = new C1.Data.C1DataSet();
			this.dtMes = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.cdsProdSemana)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPronostico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPronostico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMes)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsProdSemana
			// 
			this.cdsProdSemana.BindingContextCtrl = this;
			this.cdsProdSemana.DataLibrary = "LibFacturacion";
			this.cdsProdSemana.DataLibraryUrl = "";
			this.cdsProdSemana.DataSetDef = "dsProductoMes";
			this.cdsProdSemana.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProdSemana.Name = "cdsProdSemana";
			this.cdsProdSemana.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsProdSemana.SchemaDef = null;
			this.cdsProdSemana.BeforeFill += new C1.Data.FillEventHandler(this.cdsProdSemana_BeforeFill);
			// 
			// dtFecha
			// 
			this.dtFecha.EditValue = new System.DateTime(2004, 10, 31, 0, 0, 0, 0);
			this.dtFecha.Location = new System.Drawing.Point(136, 8);
			this.dtFecha.Name = "dtFecha";
			// 
			// dtFecha.Properties
			// 
			this.dtFecha.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.TabIndex = 1;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "ProdSemana";
			this.ultraGrid1.DataSource = this.cdsProdSemana;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Artículo";
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 174;
			ultraGridColumn5.Header.Caption = "Cant. Mes";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn6.Header.Caption = "Peso Mes";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 78;
			ultraGridColumn7.Header.Caption = "Sem. 1";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn8.Header.Caption = "Sem. 2";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.Caption = "Sem. 3";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.Header.Caption = "Sem. 4";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn11.Header.Caption = "Sem. 5";
			ultraGridColumn11.Header.VisiblePosition = 10;
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
																										 ultraGridColumn16});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(688, 232);
			this.ultraGrid1.TabIndex = 2;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.cdvArticulo;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraDropDown1.DisplayLayout.Appearance = appearance6;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 148;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraDropDown1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraDropDown1.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraDropDown1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraDropDown1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraDropDown1.DisplayMember = "Articulo";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(232, 184);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(264, 80);
			this.ultraDropDown1.TabIndex = 3;
			this.ultraDropDown1.Text = "ultraDropDown1";
			this.ultraDropDown1.ValueMember = "idArticulo";
			this.ultraDropDown1.Visible = false;
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(280, 8);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 4;
			this.btGenerar.Text = "&Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(120, 288);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 5;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(256, 288);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 6;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdvPronostico
			// 
			this.cdvPronostico.BindingContextCtrl = this;
			this.cdvPronostico.DataSet = this.cdsPronostico;
			this.cdvPronostico.TableName = "";
			this.cdvPronostico.TableViewName = "ProdPronostico";
			// 
			// cdsPronostico
			// 
			this.cdsPronostico.BindingContextCtrl = this;
			this.cdsPronostico.DataLibrary = "LibFacturacion";
			this.cdsPronostico.DataLibraryUrl = "";
			this.cdsPronostico.DataSetDef = "dsProdPronostico";
			this.cdsPronostico.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPronostico.Name = "cdsPronostico";
			this.cdsPronostico.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPronostico.SchemaDef = null;
			this.cdsPronostico.BeforeFill += new C1.Data.FillEventHandler(this.cdsPronostico_BeforeFill);
			// 
			// dtMes
			// 
			this.dtMes.AllowMonthSelection = true;
			this.dtMes.BorderStyleDay = Infragistics.Win.UIElementBorderStyle.None;
			dateButton1.Caption = "Hoy";
			this.dtMes.DateButtons.Add(dateButton1);
			this.dtMes.Format = "MMM/yyyy";
			this.dtMes.Location = new System.Drawing.Point(392, 8);
			this.dtMes.Name = "dtMes";
			this.dtMes.NonAutoSizeHeight = 23;
			this.dtMes.Size = new System.Drawing.Size(120, 21);
			this.dtMes.SpinButtonsVisible = true;
			this.dtMes.TabIndex = 7;
			// 
			// ProductoMes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(728, 334);
			this.Controls.Add(this.dtMes);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.dtFecha);
			this.Name = "ProductoMes";
			this.Text = "ProductoMes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ProductoMes_Closing);
			this.Load += new System.EventHandler(this.ProductoMes_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsProdSemana)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPronostico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPronostico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtMes)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			DateTime dtFechaActual = (DateTime) this.dtFecha.EditValue;
		}

		private void ProductoMes_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Producto Mes'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);

			Cursor = Cursors.Default;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Mes"].DefaultCellValue = 
				DateTime.Today.Month;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Anio"].DefaultCellValue = 
				DateTime.Today.Year;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsProdSemana.Update();
			this.Close();
		}

		private void CopiaPronostico(int IdArtAntes, double fTotal)
		{
			DateTime dtFechaMes = (DateTime) this.dtFecha.EditValue;
			if (fTotal != 0)
			{
				C1.Data.C1DataRow drPS = this.cdsProdSemana.TableViews["ProdSemana"].AddNew();
				drPS["Mes"] = (int) dtFechaMes.Month;
				drPS["Anio"] = (int) dtFechaMes.Year;
				drPS["idArticulo"] = IdArtAntes;
				drPS["CantidadMes"] = fTotal;
				drPS["SemanaC1"] = fTotal/4;
				drPS["SemanaC2"] = fTotal/4;
				drPS["SemanaC3"] = fTotal/4;
				drPS["SemanaC4"] = fTotal/4;
				
				// Busqueda de Articulo para Peso
				C1.Data.C1DataRow drArt = 
					this.cdsArticulo.TableViews["Articulo"].Rows.Find(IdArtAntes);
				if (drArt != null)
					drPS["PesoMes"] = (double) drArt["Peso"] * fTotal;
			}
		}
		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaMes = (DateTime) this.dtFecha.EditValue;
			string stMes = "Mes" + dtFechaMes.Month.ToString().Trim();
			DataTable drPron = this.cdsPronostico.StorageDataSet.Tables["ProdPronostico"];
			int IdArtAntes = 0;
			DataRow[] drPronC = drPron.Select("", "idArticulo");
			if (drPronC.Length > 0) IdArtAntes = (int) drPronC[0]["idArticulo"];
			double fTotal = 0F;

			// Para Borrar Pronostico de Mes Anterior
			this.cdsProdSemana.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection connection = this.cdsProdSemana.Schema.Connections[0];
			SqlCommand cmdBorrar = new SqlCommand (
				"DELETE ProdSemana WHERE Mes = " + dtFechaMes.Month.ToString() + 
				" AND Anio = " + dtFechaMes.Year.ToString(), (SqlConnection)connection.DbConnection);

			try
			{
				cmdBorrar.Transaction = (SqlTransaction)
					connection.DbTransaction;
				cmdBorrar.ExecuteNonQuery();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Información");
			}
			this.cdsProdSemana.Clear();
			this.cdsProdSemana.Fill();
			// Fin de borrar pronostico

			foreach(DataRow dr in drPronC)
			{
				if ((int) dr["idArticulo"] == IdArtAntes) 
					fTotal += (double) dr[stMes];
				else 
				{
					this.CopiaPronostico(IdArtAntes, fTotal);

					fTotal = (double) dr[stMes];
					IdArtAntes = (int) dr["idArticulo"];
				}		
			}
			// Fin de Lazo de Suma para guardar ultimo articulo
			this.CopiaPronostico(IdArtAntes, fTotal);			
		}

		private void cdsProdSemana_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProdSemana.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsPronostico_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPronostico.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ProductoMes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Producto Mes'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}
	}
}
