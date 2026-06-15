using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_PreguntasValidacion.
	/// </summary>
	public class frmCre_PreguntasValidacion : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHaber;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDebe;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int idCre_DatosGeneralesV = 0;
		int iIntentos = 0;
		int iEstadoValidacion = 0;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbResultado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;

		public int iEstado = 1;

		public frmCre_PreguntasValidacion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		public frmCre_PreguntasValidacion(int IdCre_DatosGeneralesV, int IIntentos, int IEstadoValidacion )
		{

			InitializeComponent();

			idCre_DatosGeneralesV = IdCre_DatosGeneralesV;
			iIntentos = IIntentos;
			iEstadoValidacion = IEstadoValidacion;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_ResultadoPreguntas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_PreguntasValidacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pregunta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Respuesta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Intento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResultado", -1, "cmbResultado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_ResultadoPreguntas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_PreguntasValidacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pregunta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Respuesta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Intento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResultado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResultado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			this.txtHaber = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDebe = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbResultado = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResultado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// txtHaber
			// 
			this.txtHaber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHaber.Appearance = appearance1;
			this.txtHaber.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHaber.Enabled = false;
			this.txtHaber.FormatString = "#,##0.00";
			this.txtHaber.Location = new System.Drawing.Point(288, 256);
			this.txtHaber.Name = "txtHaber";
			this.txtHaber.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHaber.PromptChar = ' ';
			this.txtHaber.Size = new System.Drawing.Size(90, 21);
			this.txtHaber.TabIndex = 452;
			this.txtHaber.Visible = false;
			// 
			// txtDebe
			// 
			this.txtDebe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDebe.Appearance = appearance2;
			this.txtDebe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDebe.Enabled = false;
			this.txtDebe.FormatString = "#,##0.00";
			this.txtDebe.Location = new System.Drawing.Point(192, 256);
			this.txtDebe.Name = "txtDebe";
			this.txtDebe.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDebe.PromptChar = ' ';
			this.txtDebe.Size = new System.Drawing.Size(90, 21);
			this.txtDebe.TabIndex = 451;
			this.txtDebe.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 240);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(992, 8);
			this.groupBox1.TabIndex = 450;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Location = new System.Drawing.Point(88, 256);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 24);
			this.btnCancelar.TabIndex = 449;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Location = new System.Drawing.Point(8, 256);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(75, 24);
			this.btnAceptar.TabIndex = 448;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource4;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance3;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 176;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 215;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 235;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 540;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 214;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 161;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 137;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn8.Width = 233;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance5;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(0, -8);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1008, 240);
			this.uGridDetalle.TabIndex = 447;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
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
			// cmbResultado
			// 
			this.cmbResultado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResultado.DataSource = this.ultraDataSource1;
			this.cmbResultado.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 102;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 205;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.SummaryFooterCaption = "Grand Summaries";
			this.cmbResultado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbResultado.DisplayMember = "Resultado";
			this.cmbResultado.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbResultado.Location = new System.Drawing.Point(552, 200);
			this.cmbResultado.Name = "cmbResultado";
			this.cmbResultado.Size = new System.Drawing.Size(224, 72);
			this.cmbResultado.TabIndex = 795;
			this.cmbResultado.ValueMember = "idResultado";
			this.cmbResultado.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// frmCre_PreguntasValidacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1008, 286);
			this.ControlBox = false;
			this.Controls.Add(this.cmbResultado);
			this.Controls.Add(this.txtHaber);
			this.Controls.Add(this.txtDebe);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.uGridDetalle);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCre_PreguntasValidacion";
			this.Text = "Preguntas Validación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_PreguntasValidacion_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_PreguntasValidacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResultado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCre_PreguntasValidacion_Load(object sender, System.EventArgs e)
		{

			this.cmbResultado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Cre_EstadosPreguntasValidacion ");

			string sSQLInf = string.Format("Exec Cre_ListaPreguntasIdValidator {0},{1},{2}",idCre_DatosGeneralesV, iIntentos, iEstadoValidacion );
			SqlDataReader drInf = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
			while (drInf.Read())
			{							
				this.uGridDetalle.Rows.Band.AddNew();
				this.uGridDetalle.ActiveRow.Cells["idCre_ResultadoPreguntas"].Value = drInf.GetInt32(0);
				this.uGridDetalle.ActiveRow.Cells["idCre_PreguntasValidacion"].Value = drInf.GetInt32(1);
				this.uGridDetalle.ActiveRow.Cells["idCre_DatosGenerales"].Value = drInf.GetInt32(2);
				this.uGridDetalle.ActiveRow.Cells["Pregunta"].Value = drInf.GetString(3);
				this.uGridDetalle.ActiveRow.Cells["Respuesta"].Value = drInf.GetString(4);
				this.uGridDetalle.ActiveRow.Cells["Intento"].Value = drInf.GetInt32(5);
				this.uGridDetalle.ActiveRow.Cells["Estado"].Value = drInf.GetInt32(6);
				this.uGridDetalle.ActiveRow.Cells["idResultado"].Value = drInf.GetInt32(7);
			}
			drInf.Close();	
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmCre_PreguntasValidacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region Validacion Nulos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
			{
				if (dr.Cells["Respuesta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese una Resultado ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}	
			}
			#endregion Validacion Nulos

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPreguntas1 in this.uGridDetalle.Rows.All)
			{
				if(drPreguntas1.Cells["Respuesta"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese Respuesta ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

			}

			#region Preguntas

					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
					{
						oConexion.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConexion.CreateCommand();
						oCmdActualiza.Connection = oConexion;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 120;
								
						try
						{							
							oTransaction = oConexion.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							int iContador = 0;
                        
							foreach (Infragistics.Win.UltraWinGrid.UltraGridRow drPreguntas in this.uGridDetalle.Rows.All)
								if (drPreguntas.Cells["Respuesta"].Value.ToString() == "") iContador++;

							if (iContador > 0)
							{
								MessageBox.Show("Ingrese la respuesta del cliente en cada pregunta", "Point", MessageBoxButtons.OK);
								return;
							}

							iContador = 0;

							foreach (Infragistics.Win.UltraWinGrid.UltraGridRow drPreguntas in this.uGridDetalle.Rows.All)
								if ((int)drPreguntas.Cells["idResultado"].Value == 0) iContador++;

							if (iContador > 0)
							{
								MessageBox.Show("Seleccione si la respuesta del cliente es CORRECTA O INCORRECTA en cada pregunta", "Point", MessageBoxButtons.OK);
								return;
							}

							iContador = 0;

							foreach (Infragistics.Win.UltraWinGrid.UltraGridRow drPreguntas in this.uGridDetalle.Rows.All)
								if ((int)drPreguntas.Cells["idResultado"].Value == 2) iContador++;

							if (iContador < 2 &&
								DialogResult.Cancel == MessageBox.Show("El cliente no ha pasado la validacion\n\nContinuar con los datos ingresados?", "Point", MessageBoxButtons.OKCancel)) return;

							iIntentos++;

						

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPreguntas in this.uGridDetalle.Rows.All)
							{
								//iEstado = 3;

								iEstado = (int)drPreguntas.Cells["idResultado"].Value;

								string sSQLPreguntas = string.Format("Exec Cre_GuardaIdValidator {0}, {1}, {2},'{3}',{4},{5}",
									(int)drPreguntas.Cells["idCre_ResultadoPreguntas"].Value,
									(int)drPreguntas.Cells["idCre_PreguntasValidacion"].Value,
									(int)drPreguntas.Cells["idCre_DatosGenerales"].Value,
									drPreguntas.Cells["Respuesta"].Value.ToString(),
									iIntentos,
									iEstado); 
								oCmdActualiza.CommandText = sSQLPreguntas;
								oCmdActualiza.ExecuteNonQuery();	

							}

							oTransaction.Commit();
						}						
						catch (Exception ex)
						{
							try
							{
								oTransaction.Rollback();

								MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
							catch (Exception ex2)
							{
								MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
							}
						}
						finally
						{
							oConexion.Close();
						}
					}
	
                
			#endregion Preguntas

			DialogResult = DialogResult.OK;
		}
	}
}
