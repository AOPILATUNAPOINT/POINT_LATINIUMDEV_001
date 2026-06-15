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
	/// Descripción breve de frmNominaPruebaObservaciones.
	/// </summary>
	public class frmNominaPruebaObservaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtLocal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridObservaciones;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.PictureBox pbFoto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int idNomina = 0;
		string sCC = "";

		public frmNominaPruebaObservaciones(int IdNomina, string SCC)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			idNomina = IdNomina;
			sCC = SCC;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaPruebaObservaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaPruebaObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnObservaciones", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.label2 = new System.Windows.Forms.Label();
			this.txtLocal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridObservaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label3 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.pbFoto = new System.Windows.Forms.PictureBox();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 274);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 415;
			this.label2.Text = "Local";
			// 
			// txtLocal
			// 
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLocal.Appearance = appearance1;
			this.txtLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLocal.Enabled = false;
			this.txtLocal.Location = new System.Drawing.Point(56, 271);
			this.txtLocal.MaxLength = 10;
			this.txtLocal.Name = "txtLocal";
			this.txtLocal.Size = new System.Drawing.Size(112, 21);
			this.txtLocal.TabIndex = 414;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// uGridObservaciones
			// 
			this.uGridObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridObservaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridObservaciones.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridObservaciones.DisplayLayout.Appearance = appearance2;
			ultraGridBand1.AddButtonCaption = "SeteoF";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 400;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 90;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn5.Header.Caption = "...";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn5.Width = 20;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridObservaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridObservaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridObservaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.CellAppearance = appearance5;
			this.uGridObservaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridObservaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridObservaciones.Location = new System.Drawing.Point(176, 8);
			this.uGridObservaciones.Name = "uGridObservaciones";
			this.uGridObservaciones.Size = new System.Drawing.Size(631, 296);
			this.uGridObservaciones.TabIndex = 413;
			this.uGridObservaciones.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridObservaciones_AfterRowInsert);
			this.uGridObservaciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridObservaciones_KeyPress);
			this.uGridObservaciones.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.uGridObservaciones_BeforeRowInsert);
			this.uGridObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridObservaciones_KeyDown);
			this.uGridObservaciones.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridObservaciones_ClickCellButton);
			this.uGridObservaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridObservaciones_InitializeLayout);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 242);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 412;
			this.label3.Text = "Celular";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 209);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(31, 16);
			this.label7.TabIndex = 411;
			this.label7.Text = "Casa";
			// 
			// txtCelular
			// 
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance10;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(56, 239);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(112, 21);
			this.txtCelular.TabIndex = 410;
			// 
			// txtTelefono
			// 
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance11;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(56, 206);
			this.txtTelefono.MaxLength = 9;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 21);
			this.txtTelefono.TabIndex = 409;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 310);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(818, 8);
			this.groupBox1.TabIndex = 408;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(86, 321);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 406;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(6, 321);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 405;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// pbFoto
			// 
			this.pbFoto.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pbFoto.Location = new System.Drawing.Point(8, 7);
			this.pbFoto.Name = "pbFoto";
			this.pbFoto.Size = new System.Drawing.Size(160, 192);
			this.pbFoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pbFoto.TabIndex = 407;
			this.pbFoto.TabStop = false;
			// 
			// frmNominaPruebaObservaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(818, 350);
			this.ControlBox = false;
			this.Controls.Add(this.uGridObservaciones);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.pbFoto);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtLocal);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmNominaPruebaObservaciones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Observaciones Personal de Prueba";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNominaPruebaObservaciones_KeyDown);
			this.Load += new System.EventHandler(this.frmNominaPruebaObservaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmNominaPruebaObservaciones_Load(object sender, System.EventArgs e)
		{
			try
			{
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

				this.uGridObservaciones.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Select idNominaPruebaObservaciones, Fecha, Observaciones, Usuario From NominaPruebaObservaciones Where idNomina = {0} Order By Fecha Desc", idNomina));

				this.txtLocal.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Telefono, '') From Bodega Where idProyecto In (Select idProyecto From Proyecto Where Nombre = '{0}')", sCC));	
			}
			catch (Exception ex2)
			{
				MessageBox.Show(string.Format("Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridObservaciones.Rows.All)
			{
				if ((int)dr.Cells["idNominaPruebaObservaciones"].Value == 0)
				{
					string sSQL = string.Format("Insert Into NominaPruebaObservaciones (idNomina, Observaciones) Values ({0}, '{1}')", 
						idNomina, dr.Cells["Observaciones"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}
			}

			this.DialogResult = DialogResult.OK;
		}

		private void uGridObservaciones_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
			
		}

		private void uGridObservaciones_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idNominaPruebaObservaciones"].Value = 0;
		}

		private void uGridObservaciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.uGridObservaciones.ActiveRow.Cells["idNominaPruebaObservaciones"].Value > 0)
				if (this.uGridObservaciones.ActiveCell.Column.Key.ToString().ToUpper() == "OBSERVACIONES") e.Handled = true;
		}

		private void frmNominaPruebaObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void uGridObservaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridObservaciones);
		}

		private void uGridObservaciones_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnObservaciones")	
				FuncionesProcedimientos.AbreVistaMensajes(e.Cell.Row.Cells["Observaciones"].Value.ToString());
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
	}
}
