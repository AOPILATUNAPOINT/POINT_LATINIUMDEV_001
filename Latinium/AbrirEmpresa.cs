using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.IO.IsolatedStorage;
using System.Drawing.Printing;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Inicio.
	/// </summary>
	public class AbrirEmpresa : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsUsuario;
		private C1.Data.C1DataView cdvUsuario;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.HelpProvider helpProvider;
		private System.Windows.Forms.Label label3;
		private C1.Data.C1DataSet cdsEmpresa;
		private C1.Data.C1DataView cdvEmpresa;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresa;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private System.ComponentModel.IContainer components=null;

		public AbrirEmpresa()
		{
			InitializeComponent();
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Empresa", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contabilidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numeracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Roles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reportes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Servidor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fox");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Usuario", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iniciales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			this.cdvUsuario = new C1.Data.C1DataView();
			this.cdsUsuario = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.helpProvider = new System.Windows.Forms.HelpProvider();
			this.cmbEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvEmpresa = new C1.Data.C1DataView();
			this.cdsEmpresa = new C1.Data.C1DataSet();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvUsuario
			// 
			this.cdvUsuario.BindingContextCtrl = this;
			this.cdvUsuario.DataSet = this.cdsUsuario;
			this.cdvUsuario.Sort = "Nombre";
			this.cdvUsuario.TableName = "";
			this.cdvUsuario.TableViewName = "Usuario";
			// 
			// cdsUsuario
			// 
			this.cdsUsuario.BindingContextCtrl = this;
			this.cdsUsuario.CaseSensitive = false;
			this.cdsUsuario.DataLibrary = "LibEmpresa";
			this.cdsUsuario.DataLibraryUrl = "";
			this.cdsUsuario.DataSetDef = "dsUsuario";
			this.cdsUsuario.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUsuario.Name = "cdsUsuario";
			this.cdsUsuario.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsUsuario.SchemaDef = null;
			this.cdsUsuario.BeforeFill += new C1.Data.FillEventHandler(this.cdsUsuario_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 34);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Nombre:";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// helpProvider
			// 
			this.helpProvider.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// cmbEmpresa
			// 
			this.cmbEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresa.DataSource = this.cdvEmpresa;
			this.cmbEmpresa.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 157;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 58;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 185;
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
																										 ultraGridColumn13});
			this.cmbEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEmpresa.DisplayMember = "Empresa";
			this.cmbEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.helpProvider.SetHelpString(this.cmbEmpresa, "Seleccione la Empresa con la que desea Trabajar");
			this.cmbEmpresa.Location = new System.Drawing.Point(96, 69);
			this.cmbEmpresa.Name = "cmbEmpresa";
			this.helpProvider.SetShowHelp(this.cmbEmpresa, true);
			this.cmbEmpresa.Size = new System.Drawing.Size(176, 21);
			this.cmbEmpresa.TabIndex = 9;
			this.cmbEmpresa.ValueMember = "idEmpresa";
			// 
			// cdvEmpresa
			// 
			this.cdvEmpresa.BindingContextCtrl = this;
			this.cdvEmpresa.DataSet = this.cdsEmpresa;
			this.cdvEmpresa.RowFilter = "idTipoDB <= 10";
			this.cdvEmpresa.Sort = "Empresa";
			this.cdvEmpresa.TableName = "";
			this.cdvEmpresa.TableViewName = "Empresa";
			// 
			// cdsEmpresa
			// 
			this.cdsEmpresa.BindingContextCtrl = this;
			this.cdsEmpresa.CaseSensitive = false;
			this.cdsEmpresa.DataLibrary = "LibEmpresa";
			this.cdsEmpresa.DataLibraryUrl = "";
			this.cdsEmpresa.DataSetDef = "dsEmpresa";
			this.cdsEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEmpresa.Name = "cdsEmpresa";
			this.cdsEmpresa.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsEmpresa.SchemaDef = null;
			this.cdsEmpresa.BeforeFill += new C1.Data.FillEventHandler(this.cdsEmpresa_BeforeFill);
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataSource = this.cdvUsuario;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Empresa";
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 122;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 69;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 75;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 118;
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn18.Width = 157;
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 72;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.helpProvider.SetHelpString(this.cmbNombre, "Seleccione el usuario para Ingresar");
			this.cmbNombre.Location = new System.Drawing.Point(96, 26);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.ReadOnly = true;
			this.helpProvider.SetShowHelp(this.cmbNombre, true);
			this.cmbNombre.Size = new System.Drawing.Size(176, 21);
			this.cmbNombre.TabIndex = 10;
			this.cmbNombre.ValueMember = "idUsuario";
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.helpProvider.SetHelpString(this.btCancelar, "Cierra el Programa");
			this.btCancelar.Location = new System.Drawing.Point(184, 146);
			this.btCancelar.Name = "btCancelar";
			this.helpProvider.SetShowHelp(this.btCancelar, true);
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 12;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			this.btCancelar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCancelar_MouseUp);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 69);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(51, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Empresa:";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(32, 146);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 11;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// AbrirEmpresa
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(288, 196);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.cmbEmpresa);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.HelpButton = true;
			this.helpProvider.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.helpProvider.SetHelpString(this, "Adminstrador y Contraseña en blanco");
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "AbrirEmpresa";
			this.helpProvider.SetShowHelp(this, true);
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Abrir Empresa";
			this.Load += new System.EventHandler(this.Inicio_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		public string stNombreEmpresa = "INFOELECT - LATINIUM";
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			stNombreEmpresa = cmbEmpresa.Text;
			int IdEmpresa = (int) cmbEmpresa.Value;

			#region Directorios
			DataRow[] drDir = cdsEmpresa.StorageDataSet.Tables["Empresa"].Select(
				"idEmpresa = " + IdEmpresa.ToString());

			string stDir = (string) drDir[0]["Facturacion"];
			stDir = stDir.Trim();
			if (stDir.Length == 0)
			{
				MessageBox.Show("Directorio de Facturacion es obligatorio que exista", 
					"Información");
				return;
			}
			MenuLatinium.stDirReportes = Application.ExecutablePath + "\\";
			DialogResult = DialogResult.OK;

			if (drDir[0]["Reportes"] != DBNull.Value)
				MenuLatinium.stDirReportes = (string) drDir[0]["Reportes"];

			if (! MenuLatinium.stDirReportes.Trim().EndsWith("\\")) 
				MenuLatinium.stDirReportes = MenuLatinium.stDirReportes.Trim() + "\\";

			MenuLatinium.stDirFacturacion = stDir;
			MenuLatinium.stDirConta = stDir;
			MenuLatinium.stDirServidor = (string) drDir[0]["Servidor"];
			MenuLatinium.stDirNumero = stDir;
			MenuLatinium.stDirRoles = stDir;
			MenuLatinium.stDirRetenc = stDir;

			if (drDir[0]["idTipoDB"] != System.DBNull.Value)
				MenuLatinium.iTipoDB = (int) drDir[0]["idTipoDB"];
			else
				MenuLatinium.iTipoDB = 3;
			#endregion Directorios

			#region Funeraria
//			MenuLatinium.bFox=false;
//			MenuLatinium.sRuta="";
//			try
//			{
//				if (drDir[0]["Fox"] != System.DBNull.Value)
//					MenuLatinium.bFox = (bool) drDir[0]["Fox"];
//
//				if (drDir[0]["ruta"] != System.DBNull.Value)
//					MenuLatinium.sRuta = (string) drDir[0]["Ruta"];
//			}
//			catch{}
			#endregion Funeraria

			Close();
		}

		private void Inicio_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			DataRow[] drSeg = this.cdsEmpresa.StorageDataSet.Tables["Empresa"].Select
				("Principal = true And idTipoDB <= 10");
			if (drSeg.Length > 0)
				this.cmbEmpresa.Value = (int) drSeg[0]["idEmpresa"];
			 this.cmbNombre.Value = MenuLatinium.IdUsuario;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsUsuario_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsUsuario.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirInicio + "\\Empresa.Mdb", 1);
		}

		private void cdsEmpresa_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsEmpresa.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirInicio + "\\Empresa.Mdb", 1);
		}

		private void btCancelar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				Funcion.CreaFormulario(this);
		}

	}
}
