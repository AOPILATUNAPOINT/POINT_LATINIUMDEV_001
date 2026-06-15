using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmSolicitudSoporte.
	/// </summary>
	public class frmSolicitudSoporte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public DevExpress.XtraEditors.SimpleButton btCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkUsuarios;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProyectos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCambioInformacion;
		public DevExpress.XtraEditors.SimpleButton btImprimir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmSolicitudSoporte()
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
			this.chkUsuarios = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProyectos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCambioInformacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// chkUsuarios
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkUsuarios.Appearance = appearance1;
			this.chkUsuarios.BackColor = System.Drawing.Color.Transparent;
			this.chkUsuarios.Checked = true;
			this.chkUsuarios.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkUsuarios.Location = new System.Drawing.Point(16, 112);
			this.chkUsuarios.Name = "chkUsuarios";
			this.chkUsuarios.Size = new System.Drawing.Size(224, 22);
			this.chkUsuarios.TabIndex = 149;
			this.chkUsuarios.Text = "Solicitud Creación, Bloqueo, Accesos de Usuarios";
			// 
			// chkProyectos
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProyectos.Appearance = appearance2;
			this.chkProyectos.BackColor = System.Drawing.Color.Transparent;
			this.chkProyectos.Checked = true;
			this.chkProyectos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkProyectos.Location = new System.Drawing.Point(16, 64);
			this.chkProyectos.Name = "chkProyectos";
			this.chkProyectos.Size = new System.Drawing.Size(224, 22);
			this.chkProyectos.TabIndex = 148;
			this.chkProyectos.Text = "Solicitud Cambios o Creación de Proyectos en el Sistema";
			// 
			// chkCambioInformacion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCambioInformacion.Appearance = appearance3;
			this.chkCambioInformacion.BackColor = System.Drawing.Color.Transparent;
			this.chkCambioInformacion.Checked = true;
			this.chkCambioInformacion.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCambioInformacion.Location = new System.Drawing.Point(16, 16);
			this.chkCambioInformacion.Name = "chkCambioInformacion";
			this.chkCambioInformacion.Size = new System.Drawing.Size(224, 22);
			this.chkCambioInformacion.TabIndex = 147;
			this.chkCambioInformacion.Text = "Solicitud Cambio Informacion del Sistema";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(248, 8);
			this.groupBox1.TabIndex = 172;
			this.groupBox1.TabStop = false;
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btCancelar.Location = new System.Drawing.Point(137, 160);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 24);
			this.btCancelar.TabIndex = 174;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btImprimir.Location = new System.Drawing.Point(25, 160);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(88, 24);
			this.btImprimir.TabIndex = 173;
			this.btImprimir.Text = "&Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btGrabar_Click);
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
			// frmSolicitudSoporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(250, 198);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkUsuarios);
			this.Controls.Add(this.chkProyectos);
			this.Controls.Add(this.chkCambioInformacion);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmSolicitudSoporte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Solicitudes";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSolicitudSoporte_KeyDown);
			this.Load += new System.EventHandler(this.frmSolicitudSoporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			string sUsuario = MenuLatinium.stUsuario;
			if (this.chkCambioInformacion.Checked)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldId = new ParameterField ();
				ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
				paramFieldId.ParameterFieldName = "@idSeteo";
				ValIdId.Value = int.Parse("1");
				paramFieldId.CurrentValues.Add (ValIdId);
				paramFields.Add (paramFieldId);

				ParameterField ParamFieldUsuario = new ParameterField();
				ParameterDiscreteValue ValUsuario = new ParameterDiscreteValue();
				ParamFieldUsuario.ParameterFieldName = "@sUsuario";
				ValUsuario.Value = sUsuario;
				ParamFieldUsuario.CurrentValues.Add(ValUsuario);
				paramFields.Add(ParamFieldUsuario);

				Reporte miRep = new Reporte("SolicitudCambioInformacionSistema.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Solicitud Informacion");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
			
			if (this.chkProyectos.Checked)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldId = new ParameterField ();
				ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
				paramFieldId.ParameterFieldName = "@idSeteo";
				ValIdId.Value = int.Parse("1");
				paramFieldId.CurrentValues.Add (ValIdId);
				paramFields.Add (paramFieldId);

				ParameterField ParamFieldUsuario = new ParameterField();
				ParameterDiscreteValue ValUsuario = new ParameterDiscreteValue();
				ParamFieldUsuario.ParameterFieldName = "@sUsuario";
				ValUsuario.Value = sUsuario;
				ParamFieldUsuario.CurrentValues.Add(ValUsuario);
				paramFields.Add(ParamFieldUsuario);

				Reporte miRep = new Reporte("SolicitudCambiosCreaciónProyectos.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Solicitud Informacion");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();	
			}

			if (this.chkUsuarios.Checked)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldId = new ParameterField ();
				ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
				paramFieldId.ParameterFieldName = "@idSeteo";
				ValIdId.Value = int.Parse("1");
				paramFieldId.CurrentValues.Add (ValIdId);
				paramFields.Add (paramFieldId);

				ParameterField ParamFieldUsuario = new ParameterField();
				ParameterDiscreteValue ValUsuario = new ParameterDiscreteValue();
				ParamFieldUsuario.ParameterFieldName = "@sUsuario";
				ValUsuario.Value = sUsuario;
				ParamFieldUsuario.CurrentValues.Add(ValUsuario);
				paramFields.Add(ParamFieldUsuario);

				Reporte miRep = new Reporte("SolicitudUsuarios.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Solicitud Informacion");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmSolicitudSoporte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmSolicitudSoporte_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2010");

			if (!Funcion.Permiso("1575", cdsSeteoF))
			{					
				MessageBox.Show("No tiene acceso a imprimir Solicitudes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		}
	}
}
