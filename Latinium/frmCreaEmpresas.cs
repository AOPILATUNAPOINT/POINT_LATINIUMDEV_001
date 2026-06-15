using System;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCreaEmpresas.
	/// </summary>
	public class frmCreaEmpresas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreBaseDeDatos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreEmpresa;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPlanDeCuentas;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConsulta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCreaEmpresas()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Clave");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			this.label3 = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtNombreBaseDeDatos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombreEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.chkPlanDeCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConsulta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreBaseDeDatos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 42);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(120, 17);
			this.label3.TabIndex = 317;
			this.label3.Text = "Nombre Base De Datos";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 10);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(90, 17);
			this.lblCliente.TabIndex = 316;
			this.lblCliente.Text = "Nombre Empresa";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombreBaseDeDatos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreBaseDeDatos.Appearance = appearance1;
			this.txtNombreBaseDeDatos.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreBaseDeDatos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreBaseDeDatos.Location = new System.Drawing.Point(144, 40);
			this.txtNombreBaseDeDatos.MaxLength = 30;
			this.txtNombreBaseDeDatos.Name = "txtNombreBaseDeDatos";
			this.txtNombreBaseDeDatos.Size = new System.Drawing.Size(250, 22);
			this.txtNombreBaseDeDatos.TabIndex = 1;
			this.txtNombreBaseDeDatos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombreBaseDeDatos_KeyPress);
			// 
			// txtNombreEmpresa
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreEmpresa.Appearance = appearance2;
			this.txtNombreEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreEmpresa.Location = new System.Drawing.Point(144, 8);
			this.txtNombreEmpresa.MaxLength = 50;
			this.txtNombreEmpresa.Name = "txtNombreEmpresa";
			this.txtNombreEmpresa.Size = new System.Drawing.Size(250, 22);
			this.txtNombreEmpresa.TabIndex = 0;
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(8, 472);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(72, 23);
			this.btnAceptar.TabIndex = 3;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(88, 472);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 4;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// chkPlanDeCuentas
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPlanDeCuentas.Appearance = appearance3;
			this.chkPlanDeCuentas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkPlanDeCuentas.Checked = true;
			this.chkPlanDeCuentas.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPlanDeCuentas.Location = new System.Drawing.Point(144, 72);
			this.chkPlanDeCuentas.Name = "chkPlanDeCuentas";
			this.chkPlanDeCuentas.Size = new System.Drawing.Size(160, 21);
			this.chkPlanDeCuentas.TabIndex = 2;
			this.chkPlanDeCuentas.Text = "Mantener Plan De Cuentas";
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
			// txtGrupo
			// 
			this.txtGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGrupo.Location = new System.Drawing.Point(224, 104);
			this.txtGrupo.MaskInput = "nn";
			this.txtGrupo.MaxValue = 30;
			this.txtGrupo.MinValue = 0;
			this.txtGrupo.Name = "txtGrupo";
			this.txtGrupo.PromptChar = ' ';
			this.txtGrupo.Size = new System.Drawing.Size(88, 22);
			this.txtGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtGrupo.TabIndex = 319;
			this.txtGrupo.Value = 1;
			this.txtGrupo.ValueChanged += new System.EventHandler(this.txtGrupo_ValueChanged);
			// 
			// txtConsulta
			// 
			this.txtConsulta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConsulta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConsulta.Location = new System.Drawing.Point(8, 104);
			this.txtConsulta.Name = "txtConsulta";
			this.txtConsulta.Size = new System.Drawing.Size(208, 22);
			this.txtConsulta.TabIndex = 318;
			this.txtConsulta.ValueChanged += new System.EventHandler(this.txtConsulta_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn5.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance4;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Usuario";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.Caption = "Grupo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 50;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 280;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 81;
			ultraGridColumn5.Header.Caption = "...";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 33;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			ultraGridBand1.Header.Caption = "USUARIOS";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 9F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(384, 330);
			this.ultraGrid1.TabIndex = 320;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// frmCreaEmpresas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(400, 502);
			this.ControlBox = false;
			this.Controls.Add(this.txtConsulta);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.chkPlanDeCuentas);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtNombreBaseDeDatos);
			this.Controls.Add(this.txtNombreEmpresa);
			this.Controls.Add(this.txtGrupo);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmCreaEmpresas";
			this.Text = "Crea Nuevas Empresas";
			this.Load += new System.EventHandler(this.frmCreaEmpresas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNombreBaseDeDatos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region Verificacion
			if (this.txtNombreEmpresa.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el nombre de la nueva empresa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombreEmpresa.Focus();
				return;
			}

			if (this.txtNombreBaseDeDatos.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el nombre de la base de datos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombreBaseDeDatos.Focus();
				return;
			}
			
			string strNombre = this.txtNombreBaseDeDatos.Value.ToString().Trim();
			strNombre = strNombre.Replace(" ", "");
			strNombre = strNombre.Replace("ñ", "n");
			strNombre = strNombre.Replace("&", "y");
			strNombre = strNombre.Replace("á", "a");
			strNombre = strNombre.Replace("é", "e");
			strNombre = strNombre.Replace("í", "i");
			strNombre = strNombre.Replace("ó", "o");
			strNombre = strNombre.Replace("ú", "u");
			strNombre = Regex.Replace(strNombre, "([^0-9]|[^a-zA-Z])", "");
			strNombre = strNombre.Trim();			
			#endregion Verificacion
											
			#region Borra el plan de Cuentas y Asignacion de Cuentas - Guarda Empresas En Inicio
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec GuardaEmpresasEnInicio '{0}', '{1}', {2}", 
				this.txtNombreEmpresa.Text.ToString().Trim(), this.txtNombreBaseDeDatos.Text.ToString().Trim(), this.chkPlanDeCuentas.Checked));
			#endregion Borra el plan de Cuentas y Asignacion de Cuentas - Guarda Empresas En Inicio
			
			this.Cursor = Cursors.Default;
			MessageBox.Show("Empresa creada con éxito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			this.Close();

			this.Close();
		}

		private void txtNombreBaseDeDatos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//Funcion.CamposNumericos(sender, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCreaEmpresas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "27");

			if (!Funcion.Permiso("1913", cdsSeteoF))
			{				
				MessageBox.Show("No tiene acceso a crear  nueva empresa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			Usuarios();
		}

		private void Usuarios()
		{
			string sSQL = string.Format("Exec CreaEmpresaListaUsuarios '{0}', {1}", this.txtConsulta.Text.ToString().Trim(), (int)this.txtGrupo.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);
		}

		private void txtGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			Usuarios();
		}

		private void txtConsulta_ValueChanged(object sender, System.EventArgs e)
		{
			Usuarios();
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.ultraGrid1.UpdateData();

				string sSQL = string.Format("Exec CreaEmpresaActualizaSelUsuario {0}, {1}", 
					(int)e.Cell.Row.Cells["idUsuario"].Value, (bool)e.Cell.Row.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
