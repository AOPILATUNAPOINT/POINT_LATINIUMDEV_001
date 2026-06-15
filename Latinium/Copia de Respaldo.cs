using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Respaldo.
	/// </summary>
	public class Respaldo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigen;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnFile;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataSet cdEmpresa;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEmpresa;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Respaldo()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtOrigen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.spnFile = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cdEmpresa = new C1.Data.C1DataSet();
			this.cmbEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnFile)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 160);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(127, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Recuperar Respaldo de:";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(40, 232);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(208, 232);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtOrigen
			// 
			this.txtOrigen.Location = new System.Drawing.Point(160, 160);
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.Size = new System.Drawing.Size(152, 21);
			this.txtOrigen.TabIndex = 3;
			this.txtOrigen.Text = "C:\\Latinium\\Respaldo\\";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 192);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(135, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "No. de Archivo Diferencial";
			// 
			// spnFile
			// 
			this.spnFile.Location = new System.Drawing.Point(192, 192);
			this.spnFile.MaskInput = "nnn";
			this.spnFile.Name = "spnFile";
			this.spnFile.PromptChar = ' ';
			this.spnFile.Size = new System.Drawing.Size(72, 21);
			this.spnFile.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnFile.TabIndex = 5;
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuentaCorriente";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// cdEmpresa
			// 
			this.cdEmpresa.BindingContextCtrl = this;
			this.cdEmpresa.DataLibrary = "LibEmpresa";
			this.cdEmpresa.DataLibraryUrl = "";
			this.cdEmpresa.DataSetDef = "dsEmpresa";
			this.cdEmpresa.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdEmpresa.Name = "cdEmpresa";
			this.cdEmpresa.SchemaClassName = "LibEmpresa.DataClass";
			this.cdEmpresa.SchemaDef = null;
			// 
			// cmbEmpresa
			// 
			this.cmbEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEmpresa.DataMember = "Empresa";
			this.cmbEmpresa.DataSource = this.cdEmpresa;
			ultraGridColumn1.Header.VisiblePosition = 0;
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
																										 ultraGridColumn11});
			this.cmbEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEmpresa.DisplayMember = "Empresa";
			this.cmbEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEmpresa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEmpresa.Location = new System.Drawing.Point(104, 128);
			this.cmbEmpresa.Name = "cmbEmpresa";
			this.cmbEmpresa.Size = new System.Drawing.Size(208, 21);
			this.cmbEmpresa.TabIndex = 6;
			this.cmbEmpresa.ValueMember = "Facturacion";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 128);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(53, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Empresa:";
			// 
			// ultraLabel1
			// 
			appearance1.ForeColor = System.Drawing.Color.Red;
			this.ultraLabel1.Appearance = appearance1;
			this.ultraLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraLabel1.Location = new System.Drawing.Point(16, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(304, 96);
			this.ultraLabel1.TabIndex = 8;
			this.ultraLabel1.Text = "PRECAUCION. Esta Pantalla Regresara la Base de Datos a un estado anterior. Los da" +
				"tos perdidos son IRRECUPERABLES";
			// 
			// Respaldo
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(336, 278);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbEmpresa);
			this.Controls.Add(this.spnFile);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtOrigen);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.label1);
			this.Name = "Respaldo";
			this.Text = "Recuperar Base de Datos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Respaldo_Closing);
			this.Load += new System.EventHandler(this.Respaldo_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnFile)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEmpresa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			string stDir = txtOrigen.Text.Trim();
			if (!stDir.EndsWith("\\")) stDir += "\\";
			string stBase = "LatiniumHSQL"; //MenuLatinium.stDirFacturacion.Trim(); 
			string stDataBase = cmbEmpresa.Value.ToString().Trim();
			stDir += stBase + ".dat";
			if (!File.Exists(stDir))
			{
				MessageBox.Show("El Archivo: " + stDir
					+ " NO existe. Restauracion Terminada",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Recuperar Respaldo de: "
				+ stDir + "?", "Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Question)) return;

			Cursor = Cursors.WaitCursor;
			string stSelect = "RESTORE DATABASE "
				+ stDataBase + " FROM DISK = '"	+ stDir + "'";
			Funcion.EjecutaSQL(cdsCuenta, stSelect);

			Cursor = Cursors.Default;
			MessageBox.Show("Respaldo Recuperado: " + stDir, "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void Respaldo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Respaldo'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}

		private void Respaldo_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Respaldo'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}
	}
}
