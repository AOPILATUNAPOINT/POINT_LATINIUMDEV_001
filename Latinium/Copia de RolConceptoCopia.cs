using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolConceptoCopia.
	/// </summary>
	public class RolConceptoCopia : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRolConcepto;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraButton btCopiar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupoDest;
		private Infragistics.Win.Misc.UltraButton btOrdenar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolConceptoCopia()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			this.cdsRolConcepto = new C1.Data.C1DataSet();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbGrupoDest = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btCopiar = new Infragistics.Win.Misc.UltraButton();
			this.btOrdenar = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoDest)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsRolConcepto
			// 
			this.cdsRolConcepto.BindingContextCtrl = this;
			this.cdsRolConcepto.DataLibrary = "LibPersonal";
			this.cdsRolConcepto.DataLibraryUrl = "";
			this.cdsRolConcepto.DataSetDef = "dsRolGrupo";
			this.cdsRolConcepto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolConcepto.Name = "cdsRolConcepto";
			this.cdsRolConcepto.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolConcepto.SchemaDef = null;
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(40, 32);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(73, 14);
			this.ultraLabel5.TabIndex = 16;
			this.ultraLabel5.Text = "Grupo Origen";
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataMember = "RolGrupo";
			this.cmbGrupo.DataSource = this.cdsRolConcepto;
			this.cmbGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RolTipo";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 85;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(144, 32);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(104, 21);
			this.cmbGrupo.TabIndex = 15;
			this.cmbGrupo.ValueMember = "idRolGrupo";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(40, 64);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(77, 14);
			this.ultraLabel1.TabIndex = 18;
			this.ultraLabel1.Text = "Grupo Destino";
			// 
			// cmbGrupoDest
			// 
			this.cmbGrupoDest.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupoDest.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupoDest.DataMember = "RolGrupo";
			this.cmbGrupoDest.DataSource = this.cdsRolConcepto;
			this.cmbGrupoDest.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 85;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbGrupoDest.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGrupoDest.DisplayMember = "Grupo";
			this.cmbGrupoDest.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupoDest.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupoDest.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupoDest.Location = new System.Drawing.Point(144, 64);
			this.cmbGrupoDest.Name = "cmbGrupoDest";
			this.cmbGrupoDest.Size = new System.Drawing.Size(104, 21);
			this.cmbGrupoDest.TabIndex = 17;
			this.cmbGrupoDest.ValueMember = "idRolGrupo";
			// 
			// btCopiar
			// 
			this.btCopiar.Location = new System.Drawing.Point(48, 152);
			this.btCopiar.Name = "btCopiar";
			this.btCopiar.TabIndex = 19;
			this.btCopiar.Text = "&Copiar";
			this.btCopiar.Click += new System.EventHandler(this.btCopiar_Click);
			// 
			// btOrdenar
			// 
			this.btOrdenar.Location = new System.Drawing.Point(168, 152);
			this.btOrdenar.Name = "btOrdenar";
			this.btOrdenar.TabIndex = 20;
			this.btOrdenar.Text = "&Ordenar";
			this.btOrdenar.Click += new System.EventHandler(this.btOrdenar_Click);
			// 
			// RolConceptoCopia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 206);
			this.Controls.Add(this.btOrdenar);
			this.Controls.Add(this.btCopiar);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbGrupoDest);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.cmbGrupo);
			this.Name = "RolConceptoCopia";
			this.Text = "RolConceptoCopia";
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupoDest)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCopiar_Click(object sender, System.EventArgs e)
		{
			int iOrigen = (int) cmbGrupo.Value;
			int iDestino = (int) cmbGrupoDest.Value;
			string stExec = "Exec RolConceptoCopia " + iOrigen.ToString()
        + ", " + iDestino.ToString();
			Funcion.EjecutaSQL(cdsRolConcepto, stExec, true);
			MessageBox.Show("Rubros Copiados", "Información");
		}

		private void btOrdenar_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec RolConceptoOrdenar ";
			Funcion.EjecutaSQL(cdsRolConcepto, stExec, true);
			MessageBox.Show("Rubros Ordenados", "Información");
		}
	}
}
