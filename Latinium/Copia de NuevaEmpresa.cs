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
	/// Descripción breve de NuevaEmpresa.
	/// </summary>
	public class NuevaEmpresa : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPlanCuentas;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupoPred;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsReporte;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEmpPredef;
		private C1.Data.C1DataSet cdsEmpresa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpresa;
		private C1.Data.C1DataSet cdsTipoBase;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoBase;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsOrigen;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAsignaCuentas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public NuevaEmpresa()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TipoBaseDatos", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoBaseDatos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.chkPlanCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGrupoPred = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAsignaCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.chkEmpPredef = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.txtEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsTipoBase = new C1.Data.C1DataSet();
			this.cmbTipoBase = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.cdsOrigen = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoBase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoBase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrigen)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(108, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Nombre de Empresa";
			// 
			// chkPlanCuentas
			// 
			this.chkPlanCuentas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkPlanCuentas.Checked = true;
			this.chkPlanCuentas.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPlanCuentas.Location = new System.Drawing.Point(32, 96);
			this.chkPlanCuentas.Name = "chkPlanCuentas";
			this.chkPlanCuentas.Size = new System.Drawing.Size(192, 20);
			this.chkPlanCuentas.TabIndex = 2;
			this.chkPlanCuentas.Text = "Plan de Cuentas predefinido";
			this.chkPlanCuentas.CheckedChanged += new System.EventHandler(this.chkPlanCuentas_CheckedChanged);
			// 
			// chkGrupoPred
			// 
			this.chkGrupoPred.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkGrupoPred.Location = new System.Drawing.Point(32, 128);
			this.chkGrupoPred.Name = "chkGrupoPred";
			this.chkGrupoPred.Size = new System.Drawing.Size(216, 20);
			this.chkGrupoPred.TabIndex = 3;
			this.chkGrupoPred.Text = "Grupo de Artículos predefinido";
			// 
			// chkAsignaCuentas
			// 
			this.chkAsignaCuentas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkAsignaCuentas.Location = new System.Drawing.Point(32, 160);
			this.chkAsignaCuentas.Name = "chkAsignaCuentas";
			this.chkAsignaCuentas.Size = new System.Drawing.Size(200, 20);
			this.chkAsignaCuentas.TabIndex = 4;
			this.chkAsignaCuentas.Text = "Asignación de Cuentas";
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(56, 200);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 5;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(224, 200);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 6;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			this.cdsReporte.BeforeFill += new C1.Data.FillEventHandler(this.cdsReporte_BeforeFill);
			// 
			// chkEmpPredef
			// 
			this.chkEmpPredef.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkEmpPredef.Enabled = false;
			this.chkEmpPredef.Location = new System.Drawing.Point(216, 96);
			this.chkEmpPredef.Name = "chkEmpPredef";
			this.chkEmpPredef.Size = new System.Drawing.Size(128, 20);
			this.chkEmpPredef.TabIndex = 7;
			this.chkEmpPredef.Text = "Empresa predefinida";
			this.chkEmpPredef.Visible = false;
			// 
			// cdsEmpresa
			// 
			this.cdsEmpresa.BindingContextCtrl = this;
			this.cdsEmpresa.DataLibrary = "LibEmpresa";
			this.cdsEmpresa.DataLibraryUrl = "";
			this.cdsEmpresa.DataSetDef = "dsEmpresa";
			this.cdsEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresa.Name = "cdsEmpresa";
			this.cdsEmpresa.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsEmpresa.SchemaDef = null;
			// 
			// txtEmpresa
			// 
			this.txtEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpresa.Location = new System.Drawing.Point(152, 56);
			this.txtEmpresa.Name = "txtEmpresa";
			this.txtEmpresa.Size = new System.Drawing.Size(192, 21);
			this.txtEmpresa.TabIndex = 8;
			// 
			// cdsTipoBase
			// 
			this.cdsTipoBase.BindingContextCtrl = this;
			this.cdsTipoBase.DataLibrary = "LibEmpresa";
			this.cdsTipoBase.DataLibraryUrl = "";
			this.cdsTipoBase.DataSetDef = "dsTipoBase";
			this.cdsTipoBase.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoBase.Name = "cdsTipoBase";
			this.cdsTipoBase.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsTipoBase.SchemaDef = null;
			// 
			// cmbTipoBase
			// 
			this.cmbTipoBase.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoBase.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoBase.DataMember = "TipoBaseDatos";
			this.cmbTipoBase.DataSource = this.cdsTipoBase;
			appearance1.BackColor = System.Drawing.Color.White;
			this.cmbTipoBase.DisplayLayout.Appearance = appearance1;
			this.cmbTipoBase.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 174;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoBase.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.cmbTipoBase.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.cmbTipoBase.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbTipoBase.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbTipoBase.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.cmbTipoBase.DisplayMember = "Nombre";
			this.cmbTipoBase.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoBase.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoBase.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoBase.Location = new System.Drawing.Point(152, 24);
			this.cmbTipoBase.Name = "cmbTipoBase";
			this.cmbTipoBase.Size = new System.Drawing.Size(192, 21);
			this.cmbTipoBase.TabIndex = 9;
			this.cmbTipoBase.ValueMember = "idTipoBaseDatos";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(119, 16);
			this.label2.TabIndex = 10;
			this.label2.Text = "Tipo de Base de Datos";
			// 
			// cdsOrigen
			// 
			this.cdsOrigen.BindingContextCtrl = this;
			this.cdsOrigen.DataLibrary = "LibContabilidad";
			this.cdsOrigen.DataLibraryUrl = "";
			this.cdsOrigen.DataSetDef = "dsSeteo";
			this.cdsOrigen.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsOrigen.Name = "cdsOrigen";
			this.cdsOrigen.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsOrigen.SchemaDef = null;
			this.cdsOrigen.BeforeFill += new C1.Data.FillEventHandler(this.cdsOrigen_BeforeFill);
			// 
			// NuevaEmpresa
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(360, 246);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbTipoBase);
			this.Controls.Add(this.txtEmpresa);
			this.Controls.Add(this.chkEmpPredef);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.chkAsignaCuentas);
			this.Controls.Add(this.chkGrupoPred);
			this.Controls.Add(this.chkPlanCuentas);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "NuevaEmpresa";
			this.Text = "Nueva Empresa";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.NuevaEmpresa_Closing);
			this.Load += new System.EventHandler(this.NuevaEmpresa_Load);
			this.Leave += new System.EventHandler(this.NuevaEmpresa_Leave);
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoBase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoBase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrigen)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void chkPlanCuentas_CheckedChanged(object sender, System.EventArgs e)
		{
			this.chkAsignaCuentas.Enabled = this.chkPlanCuentas.Checked;
			if (!this.chkPlanCuentas.Checked) this.chkAsignaCuentas.Checked = false;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private bool CreaAccess()
		{
			string stDir = this.txtEmpresa.Text.Trim();
			if (stDir.Length == 0)
			{
				stDir = MenuLatinium.stDirInicio + stDir;
				this.errorProvider.SetError(this.txtEmpresa, "Ingrese Nombre de la Empresa " + stDir);
				this.txtEmpresa.Focus();
				return false;
			}
			stDir = MenuLatinium.stDirInicio + "\\" + stDir;
			string stArch = stDir + "\\Latinium.Mdb";

			DirectoryInfo di = new DirectoryInfo(stDir);

			// create the directory, only if it doesn't exist
			try
			{
				if (di.Exists == false)
					di.Create();
				File.Copy(MenuLatinium.stDirInicio + "\\Bases.mdb", stArch, false);
				foreach(DataRow dr in this.cdsReporte.StorageDataSet.Tables["Reporte"].Select())
				{
					stArch = "\\" + dr["Reporte"].ToString().Trim() + ".Rpt"; // "\\Articulo.Rpt";
					File.Copy(MenuLatinium.stDirInicio + "\\Reportes" + stArch, stDir + stArch, false);
				}
			}
			catch(System.Exception ex)
			{
				this.errorProvider.SetError(this.txtEmpresa, ex.Message);
				this.txtEmpresa.Focus();
				return false;
			}

			this.Cursor = Cursors.WaitCursor;
			Console.Write("Crea Directorios\n");
			C1.Data.C1DataRow cdr = this.cdsEmpresa.TableViews["Empresa"].AddNew();
			string stEmpresa = this.txtEmpresa.Text.ToString().Trim();
			cdr["Empresa"] = stEmpresa;
			string stDirEmpresa = MenuLatinium.stDirInicio + "\\" + stEmpresa + "\\";
			cdr["Ubicacion"] = stDirEmpresa;
			cdr["Contabilidad"] = stDirEmpresa;
			cdr["Numeracion"] = stDirEmpresa;
			cdr["Retencion"] = stDirEmpresa;
			cdr["Roles"] = stDirEmpresa;
			cdr["Articulos"] = stDirEmpresa;
			this.cdsEmpresa.Update();
			this.Cursor = Cursors.WaitCursor;
			return true;
		}


		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			#region Verificacion
			this.errorProvider.SetError(this.txtEmpresa, "");

			string strNombre = this.txtEmpresa.Value.ToString().Trim();
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
			if (strNombre.Length == 0)
			{
				errorProvider.SetError(txtEmpresa, "Ingrese Nombre de la Empresa");
				return;
			}
			if (!Regex.IsMatch(strNombre.Substring(0, 1), "[a-zA-Z]"))
			{
				errorProvider.SetError(txtEmpresa, "Nombre de la Empresa tiene que empezar con una letra");
				return;
			}

			string Programa = MenuLatinium.stDirInicio + @"\Bases\";
			string DatoOrigen = Programa +"Bases.mdf";
			if (!File.Exists(DatoOrigen))
			{
				MessageBox.Show("Nueva empresa se crea desde el servidor", "Información");
				return;
			}
			string DatoDestino = Programa + strNombre + ".mdf";
			if (File.Exists(DatoDestino))
			{
				MessageBox.Show(string.Format("Base de Datos: {0}.mdf ya existe", strNombre), "Información");
				return;
			}
			#endregion Verificacion

			#region Verificacion de Empresas Existentes
			string stEmpresa = this.txtEmpresa.Text.ToString().Trim();
			if (Funcion.Seguridad)
			{
				string stExec1 = string.Format("Exec EmpresasBase.dbo.NuevaEmpresaVerifica '{0}', '{1}'",
					strNombre, stEmpresa);
				string stMensaje1 = Funcion.sEscalarSQL(cdsOrigen, stExec1);
				if (stMensaje1.Length > 0)
				{
					if (DialogResult.No == MessageBox.Show(stMensaje1, "Información",
						MessageBoxButtons.YesNo, MessageBoxIcon.Warning)) return;
					DBEmpresa miEmpresa = new DBEmpresa();
					miEmpresa.ShowDialog();
					return;
				}
			}
			#endregion Verificacion de Empresas Existentes

			#region Separa Base de Datos Bases
			this.Cursor = Cursors.WaitCursor;
			string LogOrigen = Programa +"Bases.ldf";
			string LogDestino = Programa + strNombre + ".ldf";
			string stExec = "Exec NuevaEmpresa";
			string stMensaje = "";
			try
			{
				Funcion.sEscalarSQL(cdsOrigen, stExec);
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error en separar");
				this.Cursor = Cursors.Default;
				return;
			}
//			MessageBox.Show("Exec Ejecutado");
			if (stMensaje.Length > 0)
			{
				MessageBox.Show("Bases esta: " + stMensaje + ". No se pudo separar", "Información");
				this.Cursor = Cursors.Default;
				return;
			}
			#endregion Separa Base de Datos Bases

			#region Copia Base de Datos Bases a Destino
			try 
			{
				File.Copy(DatoOrigen, DatoDestino, false);
				File.Copy(LogOrigen, LogDestino, false);
			} 
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Copiar Base de Datos");
				this.Cursor = Cursors.Default;
				return;
			}
			#endregion Copia Base de Datos Bases a Destino

			#region Adjunta Base de Datos Bases y Destino
			stExec = "Exec NuevaEmpresa '" + Programa + "', '"
				+ strNombre + "'";
			stMensaje = Funcion.sEscalarSQL(cdsOrigen, stExec, true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show("Bases esta: " + stMensaje + ". No se pudo adjuntar", "Información");
				this.Cursor = Cursors.Default;
				return;
			}
			#endregion Separa Base de Datos Bases

			#region Borra el plan de Cuentas y Asignacion de Cuentas
			stExec = "Exec NuevaEmpresaInicio '" + strNombre + "', ";
			if (chkPlanCuentas.Checked) stExec += "1, ";
			else stExec += "0, ";
			if (chkAsignaCuentas.Checked) stExec += "1, ";
			else stExec += "0, ";
			if (chkGrupoPred.Checked) stExec += "1";
			else stExec += "0";

			stMensaje = Funcion.sEscalarSQL(cdsOrigen, stExec, true);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				this.Cursor = Cursors.Default;
				return;
			}
			#endregion Borra el plan de Cuentas y Asignacion de Cuentas

			#region Grabacion de Empresas Existentes
			if (Funcion.Seguridad || System.IO.File.Exists(@"C:\Latinium\LibAnexos1.dll"))
			{
				string stExec1 = string.Format("Exec EmpresasBase.dbo.EmpresaCrear '{0}', '{1}', '{2}', '{3}'",
					stEmpresa, MenuLatinium.stDirServidor, strNombre, MenuLatinium.stDirInicio + "\\Reportes\\");
				Funcion.EjecutaSQL(cdsOrigen, stExec1);
			}
			else
			{
				C1.Data.C1DataRow cdr = this.cdsEmpresa.TableViews[0].AddNew();
				string stDirEmpresa = MenuLatinium.stDirInicio + "\\" + strNombre + "\\";
				cdr["Empresa"] = stEmpresa;
				cdr["Servidor"] = MenuLatinium.stDirServidor;
				cdr["Reportes"] = MenuLatinium.stDirInicio + "\\Reportes\\";
				cdr["Facturacion"] = strNombre;
				cdr["Contabilidad"] = strNombre;
				cdr["Roles"] = strNombre;
				cdr["Numeracion"] = strNombre;
				cdr["Retencion"] = strNombre;
				cdr["idTipoDB"] = 3;
				this.cdsEmpresa.Update();
			}
			#endregion Verificacion de Empresas Existentes

			this.Cursor = Cursors.Default;
			MessageBox.Show("Empresa creada con éxito", "Información");
			this.Close();
		}

		private void cdsReporte_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsReporte.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirInicio + "\\Empresa.Mdb", 1);
		}

		private void NuevaEmpresa_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Nueva Empresa'";
			Funcion.EjecutaSQL(cdsOrigen, stAudita, true);
			Cursor = Cursors.Default;
			this.cmbTipoBase.Value = 3;
			this.txtEmpresa.Select();
		}


		public string stNombreAcceso = "";
		private void cdsOrigen_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			stNombreAcceso = MenuLatinium.stUsuario;
			MenuLatinium.stUsuario = "INFOELECT";
			Funcion miFuncion = new Funcion();
			cdsOrigen.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBaseWin(MenuLatinium.stDirFacturacion);
		}

		private void NuevaEmpresa_Leave(object sender, System.EventArgs e)
		{
			MenuLatinium.stUsuario = stNombreAcceso;
		}

		private void NuevaEmpresa_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Nueva Empresa'";
			Funcion.EjecutaSQL(cdsOrigen, stAudita, true);
		}

	}
}
