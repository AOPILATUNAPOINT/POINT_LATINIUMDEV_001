using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for Alarma.
	/// </summary>
	public class AlarmaMostrar : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsAlarma;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.Repository.RepositoryItemDateEdit repositoryItemDateEdit1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AlarmaMostrar()
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
			this.cdsAlarma = new C1.Data.C1DataSet();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.repositoryItemDateEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit1)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsAlarma
			// 
			this.cdsAlarma.BindingContextCtrl = this;
			this.cdsAlarma.DataLibrary = "LibFacturacion";
			this.cdsAlarma.DataLibraryUrl = "";
			this.cdsAlarma.DataSetDef = "dsAgenda";
			this.cdsAlarma.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAlarma.Name = "cdsAlarma";
			this.cdsAlarma.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsAlarma.SchemaDef = null;
			this.cdsAlarma.BeforeFill += new C1.Data.FillEventHandler(this.cdsAgenda_BeforeFill);
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridControl1.DataMember = "Alarma";
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(16, 16);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemDateEdit1});
			this.gridControl1.Size = new System.Drawing.Size(768, 352);
			this.gridControl1.TabIndex = 13;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// gridView1
			// 
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.ShowGroupPanel = false;
			// 
			// repositoryItemDateEdit1
			// 
			this.repositoryItemDateEdit1.AutoHeight = false;
			this.repositoryItemDateEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemDateEdit1.DisplayFormat.FormatString = "dd/MMM/yyyy HH:mm";
			this.repositoryItemDateEdit1.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.repositoryItemDateEdit1.EditFormat.FormatString = "dd/MMM/yyyy HH:mm";
			this.repositoryItemDateEdit1.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.repositoryItemDateEdit1.Mask.EditMask = "dd/MMM/yyyy HH:mm";
			this.repositoryItemDateEdit1.Name = "repositoryItemDateEdit1";
			// 
			// AlarmaMostrar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(800, 382);
			this.Controls.Add(this.gridControl1);
			this.Name = "AlarmaMostrar";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Alarmas Activas";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.Alarma_Load);
			this.Closed += new System.EventHandler(this.AlarmaMostrar_Closed);
			((System.ComponentModel.ISupportInitialize)(this.cdsAlarma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemDateEdit1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsAgenda_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsAlarma.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void Alarma_Load(object sender, System.EventArgs e)
		{
			string stExec = "Exec AlarmaActivar '', 3";
			gridView1.Columns.Clear();
			gridControl1.DataSource = Funcion.dvProcedimiento(cdsAlarma, stExec);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		}

		private void AlarmaMostrar_Closed(object sender, System.EventArgs e)
		{
			MenuLatinium._AlarmaMostrar = null;
		}

	}
}

