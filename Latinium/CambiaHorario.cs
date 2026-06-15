using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CambiaHorario.
	/// </summary>
	public class CambiaHorario : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsRol;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btHorario;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbHorario;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CambiaHorario()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolSucursal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrigenFondos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSeguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoEmpleador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolHorario", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LunesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MartesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MiercolesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JuevesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ViernesOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SabadoOut2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoIn");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoOut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoIn2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DomingoOut2");
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.btHorario = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.cdsRol = new C1.Data.C1DataSet();
			this.cmbHorario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHorario)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDepartamento
			// 
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataMember = "RolSucursal";
			this.cmbDepartamento.DataSource = this.cdsRol;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDepartamento.DisplayMember = "Sucursal";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepartamento.Location = new System.Drawing.Point(144, 24);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(192, 21);
			this.cmbDepartamento.TabIndex = 3;
			this.cmbDepartamento.ValueMember = "idRolSucursal";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(32, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Sucursal";
			// 
			// btHorario
			// 
			this.btHorario.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btHorario.Location = new System.Drawing.Point(112, 152);
			this.btHorario.Name = "btHorario";
			this.btHorario.Size = new System.Drawing.Size(120, 32);
			this.btHorario.TabIndex = 4;
			this.btHorario.Text = "Modificar Horario";
			this.btHorario.Click += new System.EventHandler(this.cmbHorario_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 88);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 7;
			this.label1.Text = "Horario ";
			// 
			// cdsRol
			// 
			this.cdsRol.BindingContextCtrl = this;
			this.cdsRol.DataLibrary = "LibPersonal";
			this.cdsRol.DataLibraryUrl = "";
			this.cdsRol.DataSetDef = "dsRolTablas";
			this.cdsRol.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRol.Name = "cdsRol";
			this.cdsRol.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRol.SchemaDef = null;
			// 
			// cmbHorario
			// 
			this.cmbHorario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbHorario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbHorario.DataMember = "RolHorario";
			this.cmbHorario.DataSource = this.cdsRol;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 12;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 13;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 14;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 16;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 17;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 18;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 19;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 20;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 21;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 22;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 23;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 24;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 25;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 26;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 27;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 28;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 29;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 30;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 31;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 32;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 33;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 34;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 35;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 36;
			ultraGridColumn46.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46});
			this.cmbHorario.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbHorario.DisplayMember = "Nombre";
			this.cmbHorario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHorario.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbHorario.Location = new System.Drawing.Point(144, 88);
			this.cmbHorario.Name = "cmbHorario";
			this.cmbHorario.Size = new System.Drawing.Size(192, 21);
			this.cmbHorario.TabIndex = 9;
			this.cmbHorario.ValueMember = "idRolHorario";
			this.cmbHorario.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbHorario_InitializeLayout);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// CambiaHorario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 214);
			this.Controls.Add(this.cmbHorario);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btHorario);
			this.Controls.Add(this.cmbDepartamento);
			this.Controls.Add(this.label2);
			this.Name = "CambiaHorario";
			this.Text = "Cambiar Horarios";
			this.Load += new System.EventHandler(this.CambiaHorario_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHorario)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void CambiaHorario_Load(object sender, System.EventArgs e)
		{
		
		}

		private void BorrarErrores()
		{
			errorProvider.SetError(cmbDepartamento, "");
			errorProvider.SetError(cmbHorario, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorrarErrores();
			if (cmbDepartamento.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbDepartamento, "Ingrese Departamento");
				bOk = false;
			}
			if (cmbHorario.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbHorario, "Ingrese Horario Actual");
				bOk = false;
			}
			return bOk;
		}

		private void cmbHorario_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string stMensaje = "Desea Cambiar los horarios de " 
				+ cmbDepartamento.Text + " al Horario " + cmbHorario.Text + "?";

			if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
																 
			string stSelect = "Update Personal Set Personal.idAsiDistGrupo = "
				+ cmbHorario.Value.ToString() + " Where Personal.idRolSucursal = "
				+ cmbDepartamento.Value;
			Funcion.EjecutaSQL(cdsRol, stSelect, true);
			MessageBox.Show("Horarios modificados", "Información");
		}

		private void cmbHorarioAnterior_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbHorario_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
