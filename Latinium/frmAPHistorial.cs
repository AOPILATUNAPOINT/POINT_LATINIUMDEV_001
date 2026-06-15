using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmAPHistorial.
	/// </summary>
	public class frmAPHistorial : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridObservaciones;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.PictureBox pbFoto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAPHistorial()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idNomina = 0;
		int idActividadesPersonal = 0;
		string sTarea = "";
		bool personal ;
		bool bseguimiento;
		

		public frmAPHistorial(int IdNomina, int IdActividadesPersonal, string STarea, bool Personal, bool bSeguimiento)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			// Si es Falso trae informacion de Actividades Locales
			InitializeComponent();

			idNomina = IdNomina; 
			idActividadesPersonal = IdActividadesPersonal;
			sTarea = STarea;
			personal = Personal;
			bseguimiento = bSeguimiento;

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHistorialRevisionActividadesPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActividadesPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnObservaciones");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHistorialRevisionActividadesPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActividadesPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnObservaciones");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.uGridObservaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.pbFoto = new System.Windows.Forms.PictureBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridObservaciones
			// 
			this.uGridObservaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridObservaciones.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridObservaciones.DisplayLayout.Appearance = appearance1;
			this.uGridObservaciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "SeteoF";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Format = "dd/MM/yyyy";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Width = 89;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 400;
			ultraGridColumn5.Header.Caption = "S/F";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 24;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 103;
			ultraGridColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn7.Width = 20;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridObservaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridObservaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridObservaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.CellAppearance = appearance4;
			this.uGridObservaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridObservaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridObservaciones.Location = new System.Drawing.Point(175, 80);
			this.uGridObservaciones.Name = "uGridObservaciones";
			this.uGridObservaciones.Size = new System.Drawing.Size(657, 224);
			this.uGridObservaciones.TabIndex = 424;
			this.uGridObservaciones.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridObservaciones_AfterRowInsert);
			this.uGridObservaciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridObservaciones_KeyPress);
			this.uGridObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridObservaciones_KeyDown);
			this.uGridObservaciones.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridObservaciones_ClickCellButton);
			this.uGridObservaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridObservaciones_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn5.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(7, 246);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 423;
			this.label3.Text = "Celular";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(7, 213);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(31, 16);
			this.label7.TabIndex = 422;
			this.label7.Text = "Casa";
			// 
			// txtCelular
			// 
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance9;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(55, 243);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(112, 21);
			this.txtCelular.TabIndex = 421;
			// 
			// txtTelefono
			// 
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance10;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(55, 210);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 21);
			this.txtTelefono.TabIndex = 420;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-1, 268);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(169, 8);
			this.groupBox1.TabIndex = 419;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(85, 286);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 417;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(5, 286);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 416;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// pbFoto
			// 
			this.pbFoto.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pbFoto.Location = new System.Drawing.Point(7, 11);
			this.pbFoto.Name = "pbFoto";
			this.pbFoto.Size = new System.Drawing.Size(160, 192);
			this.pbFoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pbFoto.TabIndex = 418;
			this.pbFoto.TabStop = false;
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
			// txtObservaciones
			// 
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance11;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(176, 8);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.ReadOnly = true;
			this.txtObservaciones.Size = new System.Drawing.Size(656, 64);
			this.txtObservaciones.TabIndex = 425;
			// 
			// frmAPHistorial
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(842, 312);
			this.ControlBox = false;
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.uGridObservaciones);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.pbFoto);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmAPHistorial";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seguimiento Actividades";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAPHistorial_KeyDown);
			this.Load += new System.EventHandler(this.frmAPHistorial_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmAPHistorial_Load(object sender, System.EventArgs e)
		{
			try
			{
				this.txtObservaciones.Text = sTarea;

				string sCedula = "";
			
				string sSQL = string.Format("Select NIdentificacion, ISNULL(Telefono, ''), ISNULL(Celular, '') From Nomina Where idNomina = {0}", idNomina);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{		
					sCedula = dr.GetString(0);
					this.txtTelefono.Text = dr.GetString(1);
					this.txtCelular.Text = dr.GetString(2);
				}
				dr.Close();

				FuncionesProcedimientos.NominaFoto(@"\\192.168.1.3\Fotosnomina\", sCedula, this.pbFoto);
				if (personal)
				{
					this.uGridObservaciones.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Select idHistorialRevisionActividadesPersonal, idActividadesPersonal, Fecha, Observaciones, Usuario, Estado From HistorialRevisionActividadesPersonal Where idActividadesPersonal = {0} Order By Fecha Desc", idActividadesPersonal));
					
				}
				else
				{
					this.uGridObservaciones.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Select idHistorialRevisionActividadesAlmacenes as idHistorialRevisionActividadesPersonal, idActividadesAlmacenes as idActividadesPersonal, Fecha, ObservacionesFin as Observaciones, Usuario, Estado From HistorialRevisionActividadesAlmacenes Where idActividadesAlmacenes = {0} Order By Fecha Desc", idActividadesPersonal));
					
				}
				if (!bseguimiento)
				{
					this.uGridObservaciones.DisplayLayout.Bands[0].Columns["Estado"].CellActivation = Activation.Disabled;
				}
			}
			catch (Exception ex2)
			{
				MessageBox.Show(string.Format("Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
			}
		}

		private void uGridObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridObservaciones);
		}

		private void uGridObservaciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.uGridObservaciones.ActiveRow.Cells["idHistorialRevisionActividadesPersonal"].Value > 0)
				if (this.uGridObservaciones.ActiveCell.Column.Key.ToString().ToUpper() == "OBSERVACIONES") e.Handled = true;
		}

		private void uGridObservaciones_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idHistorialRevisionActividadesPersonal"].Value = 0;
			e.Row.Cells["idActividadesPersonal"].Value = idActividadesPersonal;
			e.Row.Cells["Estado"].Value = false;
			e.Row.Cells["Fecha"].Value = DateTime.Today;
		}

		private void uGridObservaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridObservaciones_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnObservaciones")	
				FuncionesProcedimientos.AbreVistaMensajes(e.Cell.Row.Cells["Observaciones"].Value.ToString());
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (personal)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridObservaciones.Rows.All)
				{
					if ((int)dr.Cells["idHistorialRevisionActividadesPersonal"].Value == 0)
					{
						string sSQL = string.Format("Exec GuardaHistorialRevisionActividadesPersonal {0}, '{1}', '{2}', {3}", 
							idActividadesPersonal, Convert.ToDateTime(dr.Cells["Fecha"].Value).ToString("yyyyMMdd"), 
							dr.Cells["Observaciones"].Value.ToString(), 
							(bool)dr.Cells["Estado"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);					
					}
				}
			} 
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridObservaciones.Rows.All)
				{
					if ((int)dr.Cells["idHistorialRevisionActividadesPersonal"].Value == 0)
					{
						string sSQL = string.Format("Exec GuardaHistorialRevisionActividadesAlmacenes {0}, '{1}', '{2}', {3}", 
							idActividadesPersonal, Convert.ToDateTime(dr.Cells["Fecha"].Value).ToString("yyyyMMdd"), 
							dr.Cells["Observaciones"].Value.ToString(), 
							(bool)dr.Cells["Estado"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);					
					}
				}
			
			}
			this.DialogResult = DialogResult.OK;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmAPHistorial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}
	}
}
