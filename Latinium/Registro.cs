using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Management;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Registro.
	/// </summary>
	public class Registro : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btPedirClave;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPedirClave;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIngresarClave;
		private Infragistics.Win.Misc.UltraButton btIngresarClave;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkContabilidad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFacturacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAnexo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProduccion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImportaciones;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGerencial;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivoFijo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRolPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRelojBiometrico;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnUsuarios;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtModulos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpresa;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtModelo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtActivacion;
		private Infragistics.Win.Misc.UltraButton btActivacion;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel lblFechaLimite;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Registro()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.chkContabilidad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFacturacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAnexo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkActivoFijo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkRolPago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProduccion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkImportaciones = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGerencial = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkRelojBiometrico = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btPedirClave = new Infragistics.Win.Misc.UltraButton();
			this.txtPedirClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIngresarClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btIngresarClave = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.spnUsuarios = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtModulos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtModelo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtActivacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btActivacion = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.lblFechaLimite = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.txtPedirClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresarClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnUsuarios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtActivacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			this.SuspendLayout();
			// 
			// chkContabilidad
			// 
			this.chkContabilidad.Location = new System.Drawing.Point(40, 48);
			this.chkContabilidad.Name = "chkContabilidad";
			this.chkContabilidad.TabIndex = 0;
			this.chkContabilidad.Text = "Contabilidad";
			// 
			// chkFacturacion
			// 
			this.chkFacturacion.Location = new System.Drawing.Point(40, 72);
			this.chkFacturacion.Name = "chkFacturacion";
			this.chkFacturacion.TabIndex = 1;
			this.chkFacturacion.Text = "Facturación";
			// 
			// chkAnexo
			// 
			this.chkAnexo.Location = new System.Drawing.Point(40, 96);
			this.chkAnexo.Name = "chkAnexo";
			this.chkAnexo.Size = new System.Drawing.Size(136, 20);
			this.chkAnexo.TabIndex = 2;
			this.chkAnexo.Text = "Anexo Transaccional";
			// 
			// chkActivoFijo
			// 
			this.chkActivoFijo.Location = new System.Drawing.Point(40, 120);
			this.chkActivoFijo.Name = "chkActivoFijo";
			this.chkActivoFijo.Size = new System.Drawing.Size(136, 20);
			this.chkActivoFijo.TabIndex = 3;
			this.chkActivoFijo.Text = "Activos Fijos";
			// 
			// chkRolPago
			// 
			this.chkRolPago.Location = new System.Drawing.Point(40, 144);
			this.chkRolPago.Name = "chkRolPago";
			this.chkRolPago.Size = new System.Drawing.Size(136, 20);
			this.chkRolPago.TabIndex = 4;
			this.chkRolPago.Text = "Roles de Pago";
			// 
			// chkProduccion
			// 
			this.chkProduccion.Location = new System.Drawing.Point(232, 48);
			this.chkProduccion.Name = "chkProduccion";
			this.chkProduccion.Size = new System.Drawing.Size(152, 20);
			this.chkProduccion.TabIndex = 5;
			this.chkProduccion.Text = "Producción";
			// 
			// chkImportaciones
			// 
			this.chkImportaciones.Location = new System.Drawing.Point(232, 72);
			this.chkImportaciones.Name = "chkImportaciones";
			this.chkImportaciones.Size = new System.Drawing.Size(152, 20);
			this.chkImportaciones.TabIndex = 6;
			this.chkImportaciones.Text = "Importaciones";
			// 
			// chkGerencial
			// 
			this.chkGerencial.Location = new System.Drawing.Point(232, 96);
			this.chkGerencial.Name = "chkGerencial";
			this.chkGerencial.Size = new System.Drawing.Size(152, 20);
			this.chkGerencial.TabIndex = 7;
			this.chkGerencial.Text = "Gerencial";
			// 
			// chkRelojBiometrico
			// 
			this.chkRelojBiometrico.Location = new System.Drawing.Point(232, 120);
			this.chkRelojBiometrico.Name = "chkRelojBiometrico";
			this.chkRelojBiometrico.Size = new System.Drawing.Size(144, 20);
			this.chkRelojBiometrico.TabIndex = 8;
			this.chkRelojBiometrico.Text = "Reloj Biometrico";
			// 
			// btPedirClave
			// 
			this.btPedirClave.Location = new System.Drawing.Point(216, 192);
			this.btPedirClave.Name = "btPedirClave";
			this.btPedirClave.Size = new System.Drawing.Size(88, 23);
			this.btPedirClave.TabIndex = 10;
			this.btPedirClave.Text = "Pedir Clave";
			this.btPedirClave.Click += new System.EventHandler(this.btPedirClave_Click);
			// 
			// txtPedirClave
			// 
			appearance1.ForeColor = System.Drawing.Color.Blue;
			this.txtPedirClave.Appearance = appearance1;
			this.txtPedirClave.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPedirClave.Location = new System.Drawing.Point(312, 184);
			this.txtPedirClave.Name = "txtPedirClave";
			this.txtPedirClave.Size = new System.Drawing.Size(168, 33);
			this.txtPedirClave.TabIndex = 11;
			// 
			// txtIngresarClave
			// 
			appearance2.ForeColor = System.Drawing.Color.Green;
			this.txtIngresarClave.Appearance = appearance2;
			this.txtIngresarClave.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIngresarClave.Location = new System.Drawing.Point(312, 256);
			this.txtIngresarClave.Name = "txtIngresarClave";
			this.txtIngresarClave.Size = new System.Drawing.Size(168, 33);
			this.txtIngresarClave.TabIndex = 13;
			// 
			// btIngresarClave
			// 
			this.btIngresarClave.Location = new System.Drawing.Point(216, 224);
			this.btIngresarClave.Name = "btIngresarClave";
			this.btIngresarClave.Size = new System.Drawing.Size(88, 23);
			this.btIngresarClave.TabIndex = 12;
			this.btIngresarClave.Text = "Ingresar Clave";
			this.btIngresarClave.Click += new System.EventHandler(this.btIngresarClave_Click);
			this.btIngresarClave.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btIngresarClave_MouseUp);
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// spnUsuarios
			// 
			this.spnUsuarios.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Usuarios"));
			this.spnUsuarios.Location = new System.Drawing.Point(304, 152);
			this.spnUsuarios.MaskInput = "nnn";
			this.spnUsuarios.MaxValue = 999;
			this.spnUsuarios.MinValue = 1;
			this.spnUsuarios.Name = "spnUsuarios";
			this.spnUsuarios.PromptChar = ' ';
			this.spnUsuarios.Size = new System.Drawing.Size(72, 21);
			this.spnUsuarios.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnUsuarios.TabIndex = 15;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(232, 152);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(49, 14);
			this.ultraLabel1.TabIndex = 16;
			this.ultraLabel1.Text = "Usuarios";
			// 
			// txtModulos
			// 
			this.txtModulos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Modulos"));
			this.txtModulos.Location = new System.Drawing.Point(376, 16);
			this.txtModulos.Name = "txtModulos";
			this.txtModulos.Size = new System.Drawing.Size(56, 21);
			this.txtModulos.TabIndex = 14;
			// 
			// txtEmpresa
			// 
			this.txtEmpresa.Location = new System.Drawing.Point(160, 16);
			this.txtEmpresa.Name = "txtEmpresa";
			this.txtEmpresa.Size = new System.Drawing.Size(168, 21);
			this.txtEmpresa.TabIndex = 17;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(96, 16);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(49, 14);
			this.ultraLabel2.TabIndex = 18;
			this.ultraLabel2.Text = "Empresa";
			// 
			// txtModelo
			// 
			this.txtModelo.AcceptsReturn = true;
			this.txtModelo.AcceptsTab = true;
			this.txtModelo.AutoSize = false;
			this.txtModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtModelo.Location = new System.Drawing.Point(16, 248);
			this.txtModelo.Multiline = true;
			this.txtModelo.Name = "txtModelo";
			this.txtModelo.ReadOnly = true;
			this.txtModelo.Scrollbars = System.Windows.Forms.ScrollBars.Both;
			this.txtModelo.Size = new System.Drawing.Size(200, 120);
			this.txtModelo.TabIndex = 19;
			this.txtModelo.ValueChanged += new System.EventHandler(this.txtModelo_ValueChanged);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtActivacion
			// 
			this.txtActivacion.Location = new System.Drawing.Point(312, 224);
			this.txtActivacion.Name = "txtActivacion";
			this.txtActivacion.PasswordChar = '*';
			this.txtActivacion.Size = new System.Drawing.Size(88, 21);
			this.txtActivacion.TabIndex = 20;
			// 
			// btActivacion
			// 
			this.btActivacion.Location = new System.Drawing.Point(312, 296);
			this.btActivacion.Name = "btActivacion";
			this.btActivacion.Size = new System.Drawing.Size(88, 24);
			this.btActivacion.TabIndex = 21;
			this.btActivacion.Text = "Activación";
			this.btActivacion.Click += new System.EventHandler(this.btActivacion_Click);
			this.btActivacion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btActivacion_MouseUp);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(216, 264);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 22;
			this.label1.Text = "Clave Generada";
			// 
			// cmbFecha
			// 
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.FechaActivacion"));
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(104, 224);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 16;
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.TabIndex = 23;
			this.cmbFecha.Value = new System.DateTime(2008, 3, 21, 0, 0, 0, 0);
			this.cmbFecha.Visible = false;
			// 
			// cmbTipo
			// 
			this.cmbTipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.TipoActivacion"));
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Total";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Parcial";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Academica";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Items.Add(valueListItem2);
			this.cmbTipo.Items.Add(valueListItem3);
			this.cmbTipo.Location = new System.Drawing.Point(104, 192);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(104, 21);
			this.cmbTipo.TabIndex = 24;
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.Location = new System.Drawing.Point(24, 192);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(72, 16);
			this.ultraLabel3.TabIndex = 25;
			this.ultraLabel3.Text = "Tipo Activac.";
			// 
			// lblFechaLimite
			// 
			this.lblFechaLimite.Location = new System.Drawing.Point(24, 224);
			this.lblFechaLimite.Name = "lblFechaLimite";
			this.lblFechaLimite.Size = new System.Drawing.Size(72, 16);
			this.lblFechaLimite.TabIndex = 26;
			this.lblFechaLimite.Text = "Fecha Limite";
			this.lblFechaLimite.Visible = false;
			// 
			// Registro
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(504, 382);
			this.Controls.Add(this.lblFechaLimite);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btActivacion);
			this.Controls.Add(this.txtActivacion);
			this.Controls.Add(this.txtModelo);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtEmpresa);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.spnUsuarios);
			this.Controls.Add(this.txtModulos);
			this.Controls.Add(this.txtIngresarClave);
			this.Controls.Add(this.btIngresarClave);
			this.Controls.Add(this.txtPedirClave);
			this.Controls.Add(this.btPedirClave);
			this.Controls.Add(this.chkRelojBiometrico);
			this.Controls.Add(this.chkGerencial);
			this.Controls.Add(this.chkImportaciones);
			this.Controls.Add(this.chkProduccion);
			this.Controls.Add(this.chkRolPago);
			this.Controls.Add(this.chkActivoFijo);
			this.Controls.Add(this.chkAnexo);
			this.Controls.Add(this.chkFacturacion);
			this.Controls.Add(this.chkContabilidad);
			this.Name = "Registro";
			this.Text = "Registro";
			this.Load += new System.EventHandler(this.Registro_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtPedirClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresarClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnUsuarios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtActivacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		class HardDrive
		{
			private string model = null;
			private string type = null;
			private string serialNo = null;
			private string memoria = null;
			private string equipo = null;
			public string Equipo
			{
				get {return equipo;}
				set {equipo = value;}
			} 
			public string Memoria
			{
				get {return memoria;}
				set {memoria = value;}
			} 
			public string Modelo
			{
				get {return model;}
				set {model = value;}
			} 
			public string Tipo
			{
				get {return type;}
				set {type = value;}
			}
			public string SerialNo
			{
				get {return serialNo;}
				set {serialNo = value;}
			}
		}

		private void btPedirClave_Click(object sender, System.EventArgs e)
		{
			ArrayList hdCollection = new ArrayList();

			ManagementObjectSearcher searcher = new
				ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive");

			foreach(ManagementObject wmi_HD in searcher.Get())
			{
				HardDrive hd = new HardDrive();
				hd.Modelo = wmi_HD["Model"].ToString();
				hd.Tipo  = wmi_HD["InterfaceType"].ToString();
				hdCollection.Add(hd);
				break;
			}
			searcher = new
				ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia");
   
			int i = 0;
			string stSerie = "";
			foreach(ManagementObject wmi_HD in searcher.Get())
			{
				HardDrive hd = (HardDrive)hdCollection[i];

				if (wmi_HD["SerialNumber"] == null)
					hd.SerialNo = "Ninguno";
				else
				{
					hd.SerialNo = wmi_HD["SerialNumber"].ToString();
					stSerie = hd.SerialNo;
				}
				++i;
				break;
			}

			try
			{
				searcher = new
					ManagementObjectSearcher(@"root\CIMV2", "SELECT * FROM Win32_MemoryDevice");
   
				i = 0;
				foreach(ManagementObject wmi_HD in searcher.Get())
				{
					HardDrive hd = (HardDrive)hdCollection[i];
					hd.Memoria = wmi_HD["EndingAddress"].ToString();
					hd.Equipo = wmi_HD["SystemName"].ToString();
					break;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al leer memoria");
			}

			string stMensaje = "";
			foreach(HardDrive hd in hdCollection)
			{
				stMensaje += "Modelo\t\t: " + hd.Modelo;
				stMensaje += "\nTipo\t\t: " + hd.Tipo;
				stMensaje += "\nSerie No.\t: " + hd.SerialNo.Trim();
				stMensaje += "\nMemoria \t: " + hd.Memoria;
				stMensaje += "\nEquipo\t: " + hd.Equipo;
				stSerieDisco = hd.SerialNo.Trim();
				break;
			}
			txtModelo.Text = stMensaje;
			stSerie = stSerieDisco.ToUpper().Trim();
			stSerie = stSerie.Replace("0", "1");
			stSerie = stSerie.Replace("O", "A");
			try
			{
				stSerieDisco = stSerie.Substring(0, 8);
			}
			catch//System.Exception ex)
			{
				stSerieDisco = stSerie;
			}
			txtPedirClave.Text = stSerieDisco;
		}

		private string CreaClave(string stSerie)
		{
			string stClave1 = Verifica();
			if (stClave1.Length == 0) return "";
			string stClave = Funcion.CreaHash(stSerie + stClave1);
			return stClave.ToUpper().Substring(0, 8);
		}

		private void Memoria()
		{
			try
			{
				ManagementObjectSearcher searcher = new
					ManagementObjectSearcher(@"root\CIMV2", "SELECT * FROM Win32_MemoryDevice");
   
				foreach(ManagementObject wmi_HD in searcher.Get())
				{
					HardDrive hd = new HardDrive();
					hd.Memoria = wmi_HD["DeviceID"].ToString();
					hd.Equipo = wmi_HD["SystemName"].ToString();
					MessageBox.Show(hd.Modelo);
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al leer memoria");
			}
		}

		private int CreaModulos()
		{
			int iModulos = 0;
			if (chkActivoFijo.Checked) iModulos += 1;
			if (chkAnexo.Checked) iModulos += 2;
			if (chkContabilidad.Checked) iModulos += 4;
			if (chkFacturacion.Checked) iModulos += 8;
			if (chkGerencial.Checked) iModulos += 16;
			if (chkImportaciones.Checked) iModulos += 32;
			if (chkProduccion.Checked) iModulos += 64;
			if (chkRelojBiometrico.Checked) iModulos += 128;
			if (chkRolPago.Checked) iModulos += 256;
			return iModulos;
		}

		private void MostrarModulos(int iModulos)
		{
			if ((iModulos & 1) == 1) chkActivoFijo.Checked = true;
			if ((iModulos & 2) == 2) chkAnexo.Checked = true;
			if ((iModulos & 4) == 4) chkContabilidad.Checked = true;
			if ((iModulos & 8) == 8) chkFacturacion.Checked = true;
			if ((iModulos & 16) == 16) chkGerencial.Checked = true;
			if ((iModulos & 32) == 32) chkImportaciones.Checked = true;
			if ((iModulos & 64) == 64) chkProduccion.Checked = true;
			if ((iModulos & 128) == 128) chkRelojBiometrico.Checked = true;
			if ((iModulos & 256) == 256) chkRolPago.Checked = true;
		}

		private string stSerieDisco = "";
		private string Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(chkContabilidad, "");
			errorProvider.SetError(spnUsuarios, "");
			errorProvider.SetError(cmbTipo, "");
			errorProvider.SetError(cmbFecha, "");

			int iModulos = CreaModulos();
			if (iModulos == 0)
			{
				errorProvider.SetError(chkContabilidad, "Ingrese al menos un modulo");
				bOk = false;
			}
			int iUsuarios = spnUsuarios.Value == System.DBNull.Value ? 0 : (int) spnUsuarios.Value;
			if (iUsuarios == 0)
			{
				errorProvider.SetError(spnUsuarios, "Ingrese al menos un usuario");
				bOk = false;
			}
			int iTipo = cmbTipo.Value == DBNull.Value ? -1 : (int) cmbTipo.Value;
			string stTipo = "";
			if (iTipo == -1)
			{
				errorProvider.SetError(cmbTipo, "Ingrese tipo de Activacion");
				bOk = false;
			}
			stTipo = iTipo.ToString();
			if (iTipo == 1)
			{
				DateTime dtFecha = cmbFecha.Value == DBNull.Value ? DateTime.Today.AddMonths(-1) : (DateTime) cmbFecha.Value;
				if (dtFecha < (DateTime) DateTime.Today)
				{
					errorProvider.SetError(cmbFecha, "Ingrese fecha");
					bOk = false;
				}
				stTipo += dtFecha.ToString("ddMMyyyy");
			}
			if (bOk) return iModulos.ToString() + iUsuarios.ToString() + stTipo;
			else return "";
		}

		private void btIngresarClave_Click(object sender, System.EventArgs e)
		{
			string stClave = Funcion.CreaHash(txtActivacion.Text.Trim().ToLower());
			errorProvider.SetError(txtActivacion, "");
			errorProvider.SetError(txtPedirClave, "");
			if(stClave != Funcion.ClaveMaestra)
			{
				errorProvider.SetError(txtActivacion, "Codigo Maestro de Activacion Incorrecto");
				return;
			}

			if (Verifica().Length == 0) return;
			stSerieDisco = txtPedirClave.Text.Trim();
			if (stSerieDisco.Length < 8)
			{
				errorProvider.SetError(txtPedirClave, "Ingrese Serie de 8 caracteres");
				return;
			}
			stClave = CreaClave(stSerieDisco);
			txtIngresarClave.Text = stClave;
			int iModulos = CreaModulos();
			int iUsuarios = (int) spnUsuarios.Value;
			DateTime dtValida = cmbFecha.Value == DBNull.Value ? DateTime.Today.AddDays(-1) : (DateTime) cmbFecha.Value;

			string stAudita = "Exec AuditaCrear 56, 6, '" + stClave + "|"+ stSerieDisco+"'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			stAudita = "Exec AuditaCrear 56, 6, '" + iModulos.ToString() + "|" + iUsuarios.ToString() 
				+ "|" + cmbTipo.Value.ToString() + "|" + dtValida.ToString("dd/MMM/yyyy") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void Registro_Load(object sender, System.EventArgs e)
		{
			txtModulos.Width = 0;
			int iModulos = txtModulos.Value == DBNull.Value ? 0: (int) txtModulos.Value;
			MostrarModulos(iModulos);
			VisibleFecha();
		}


		private void btActivacion_Click(object sender, System.EventArgs e)
		{
			string stClave = txtIngresarClave.Text.Trim();
			errorProvider.SetError(txtIngresarClave, "");
			errorProvider.SetError(txtPedirClave, "");
			if (stClave.Length < 8)
			{
				errorProvider.SetError(txtIngresarClave, "Ingrese Clave de 8 caracteres");
				return;
			}
			if (txtPedirClave.Text.Trim().Length < 8)
			{
				errorProvider.SetError(txtPedirClave, "Ingrese Serie de 8 caracteres");
				return;
			}
			if (Verifica().Length == 0) return;
			string stClaveVerifica = CreaClave(txtPedirClave.Text.Trim());
			if (stClave != stClaveVerifica)
			{
				string stAuditaE = "Exec AuditaCrear 57, 6, 'E:" + stClave + "|" + stSerieDisco + "'";
				Funcion.EjecutaSQL(cdsSeteo, stAuditaE, true);
				errorProvider.SetError(txtIngresarClave, "Ingrese Clave proporcionado por INFOELECT");
				return;
			}
			int iModulos = CreaModulos();
			int iUsuarios = (int) spnUsuarios.Value;
			DateTime dtValida = cmbFecha.Value == DBNull.Value ? DateTime.Today.AddDays(-1) : (DateTime) cmbFecha.Value;
			string stSelect = "Update Seteo Set Modulos = " + iModulos + ", Usuarios = "
				+ iUsuarios.ToString() + ", SerieDisco = '" + stSerieDisco 
				+ "', ClaveUsr = '" + stClave + "', FechaActivacion = '" 
				+ dtValida.ToString("yyyyMMdd") + "', TipoActivacion = " + cmbTipo.Value.ToString();
			Funcion.EjecutaSQL(cdsSeteo, stSelect, true);
			string stAudita = "Exec AuditaCrear 57, 6, '" + stClave + "|" + stSerieDisco + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			stAudita = "Exec AuditaCrear 57, 6, '" + iModulos.ToString() + "|" + iUsuarios.ToString() 
				+ "|" + cmbTipo.Value.ToString() + "|" + dtValida.ToString("dd/MMM/yyyy") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
			MessageBox.Show("Activacion Realizada", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void VisibleFecha()
		{
			int iTipoAct = 0;
			if (cmbTipo.Value == DBNull.Value || cmbTipo.Value ==  null) cmbTipo.Value = 0;
			iTipoAct = (int) cmbTipo.Value;
			if (iTipoAct == 1)
			{
				lblFechaLimite.Visible = true;
				cmbFecha.Visible = true;
			}
			else
			{
				lblFechaLimite.Visible = false;
				cmbFecha.Visible = false;
			}
		}
		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			VisibleFecha();
		}

		private void btIngresarClave_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Reporte miRep = new Reporte("Auditoria.rpt", "{Auditoria.idTipoAudita} = 56");
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btActivacion_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Reporte miRep = new Reporte("Auditoria.rpt", "{Auditoria.idTipoAudita} = 57");
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void txtModelo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

	}
}
