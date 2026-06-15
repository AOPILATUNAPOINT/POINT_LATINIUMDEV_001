using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for RolSalarioDigno.
	/// </summary>
	public class RolSalarioDigno : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.TextEdit txtUtilidad;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.SpinEdit spnAnio;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.SimpleButton btGenera;
		private DevExpress.XtraEditors.LabelControl labelControl3;
		private DevExpress.XtraEditors.TextEdit txtSalarioDigno;
		private C1.Data.C1DataSet cdsSeguridad;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolSalarioDigno()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.txtUtilidad = new DevExpress.XtraEditors.TextEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.spnAnio = new DevExpress.XtraEditors.SpinEdit();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.btGenera = new DevExpress.XtraEditors.SimpleButton();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.txtSalarioDigno = new DevExpress.XtraEditors.TextEdit();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidad.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnAnio.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalarioDigno.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// txtUtilidad
			// 
			this.txtUtilidad.EditValue = "0";
			this.txtUtilidad.Location = new System.Drawing.Point(312, 16);
			this.txtUtilidad.Name = "txtUtilidad";
			// 
			// txtUtilidad.Properties
			// 
			this.txtUtilidad.Properties.Appearance.Options.UseTextOptions = true;
			this.txtUtilidad.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtUtilidad.Properties.DisplayFormat.FormatString = "n2";
			this.txtUtilidad.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtUtilidad.Properties.EditFormat.FormatString = "n2";
			this.txtUtilidad.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtUtilidad.Properties.Mask.BeepOnError = true;
			this.txtUtilidad.Properties.Mask.EditMask = "n2";
			this.txtUtilidad.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
			this.txtUtilidad.Size = new System.Drawing.Size(88, 20);
			this.txtUtilidad.TabIndex = 0;
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(16, 16);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(72, 14);
			this.labelControl1.TabIndex = 1;
			this.labelControl1.Text = "Salario Digno";
			// 
			// labelControl2
			// 
			this.labelControl2.Location = new System.Drawing.Point(440, 16);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(32, 14);
			this.labelControl2.TabIndex = 3;
			this.labelControl2.Text = "Año";
			// 
			// spnAnio
			// 
			this.spnAnio.EditValue = new System.Decimal(new int[] {
																															0,
																															0,
																															0,
																															0});
			this.spnAnio.Location = new System.Drawing.Point(488, 16);
			this.spnAnio.Name = "spnAnio";
			// 
			// spnAnio.Properties
			// 
			this.spnAnio.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnAnio.Size = new System.Drawing.Size(64, 20);
			this.spnAnio.TabIndex = 4;
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 48);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(600, 256);
			this.gridControl1.TabIndex = 5;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// gridView1
			// 
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			// 
			// btGenera
			// 
			this.btGenera.Location = new System.Drawing.Point(560, 16);
			this.btGenera.Name = "btGenera";
			this.btGenera.Size = new System.Drawing.Size(56, 24);
			this.btGenera.TabIndex = 6;
			this.btGenera.Text = "Genera";
			this.btGenera.Click += new System.EventHandler(this.btGenera_Click);
			// 
			// labelControl3
			// 
			this.labelControl3.Location = new System.Drawing.Point(224, 16);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(80, 14);
			this.labelControl3.TabIndex = 8;
			this.labelControl3.Text = "Utilidad Empresa";
			// 
			// txtSalarioDigno
			// 
			this.txtSalarioDigno.Location = new System.Drawing.Point(96, 16);
			this.txtSalarioDigno.Name = "txtSalarioDigno";
			// 
			// txtSalarioDigno.Properties
			// 
			this.txtSalarioDigno.Properties.Appearance.Options.UseTextOptions = true;
			this.txtSalarioDigno.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.txtSalarioDigno.Properties.DisplayFormat.FormatString = "n2";
			this.txtSalarioDigno.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtSalarioDigno.Properties.EditFormat.FormatString = "n2";
			this.txtSalarioDigno.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.txtSalarioDigno.Properties.Mask.BeepOnError = true;
			this.txtSalarioDigno.Properties.Mask.EditMask = "n2";
			this.txtSalarioDigno.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
			this.txtSalarioDigno.Size = new System.Drawing.Size(104, 20);
			this.txtSalarioDigno.TabIndex = 7;
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// RolSalarioDigno
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(632, 318);
			this.Controls.Add(this.labelControl3);
			this.Controls.Add(this.txtSalarioDigno);
			this.Controls.Add(this.btGenera);
			this.Controls.Add(this.gridControl1);
			this.Controls.Add(this.spnAnio);
			this.Controls.Add(this.labelControl2);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.txtUtilidad);
			this.Name = "RolSalarioDigno";
			this.Text = "Salario Digno";
			this.Load += new System.EventHandler(this.RolSalarioDigno_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidad.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnAnio.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalarioDigno.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void RolSalarioDigno_Load(object sender, System.EventArgs e)
		{
			spnAnio.Value = DateTime.Today.Year - 1;
			txtSalarioDigno.Text = Funcion.dEscalarSQL(cdsSeguridad, 
				"Select IsNull(SalarioDigno, 0) From SeteoRol ").ToString();
		}

		private void btGenera_Click(object sender, System.EventArgs e)
		{
			double dSalarioDigno = double.Parse(txtSalarioDigno.Text);
			string stGrabar = "Update SeteoRol Set SalarioDigno = " + dSalarioDigno;
			Funcion.EjecutaSQL(cdsSeguridad, stGrabar);
			int iAnio = (int) spnAnio.Value;
			string stExec = string.Format("Exec RolSalarioDigno {0}, {1}", iAnio, dSalarioDigno);
			gridControl1.DataSource = Funcion.dvProcedimiento(cdsSeguridad, stExec);
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
	}
}

