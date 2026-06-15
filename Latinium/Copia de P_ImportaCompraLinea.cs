using System;
using System.IO;
using System.Globalization;
using System.Data.SqlClient;
using System.Data;
using C1.C1Excel;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using System.Data.OleDb;



namespace Latinium
{
	/// <summary>
	/// Descripción breve de P_ImportaCompraLinea.
	/// </summary>
	public class P_ImportaCompraLinea : System.Windows.Forms.Form
	{

		private Infragistics.Win.Misc.UltraButton btSubir;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private C1.Data.C1DataSet cdsConciliacion;
		private Infragistics.Win.Misc.UltraButton btSubirR;
		private C1.C1Excel.C1XLBook c1XLBook;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataView cdvCompra;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.ListBox listBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkModificaExiste;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMensajeR;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMensajeC;
		private Infragistics.Win.Misc.UltraButton btLeer;
		private System.ComponentModel.IContainer components;

		public P_ImportaCompraLinea()
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
			this.btSubir = new Infragistics.Win.Misc.UltraButton();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cdsConciliacion = new C1.Data.C1DataSet();
			this.btSubirR = new Infragistics.Win.Misc.UltraButton();
			this.c1XLBook = new C1.C1Excel.C1XLBook();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdvCompra = new C1.Data.C1DataView();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.listBox1 = new System.Windows.Forms.ListBox();
			this.chkModificaExiste = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtMensajeC = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btLeer = new Infragistics.Win.Misc.UltraButton();
			this.txtMensajeR = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cdsConciliacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensajeC)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtMensajeR)).BeginInit();
			this.groupBox3.SuspendLayout();
			this.SuspendLayout();
			// 
			// btSubir
			// 
			this.btSubir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSubir.Location = new System.Drawing.Point(304, 32);
			this.btSubir.Name = "btSubir";
			this.btSubir.Size = new System.Drawing.Size(72, 24);
			this.btSubir.TabIndex = 25;
			this.btSubir.Text = "Subir";
			this.btSubir.Click += new System.EventHandler(this.btSubir_Click);
			// 
			// cdsConciliacion
			// 
			this.cdsConciliacion.BindingContextCtrl = this;
			this.cdsConciliacion.DataLibrary = "LibContabilidad";
			this.cdsConciliacion.DataLibraryUrl = "";
			this.cdsConciliacion.DataSetDef = "dsTmpConcilia";
			this.cdsConciliacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsConciliacion.Name = "cdsConciliacion";
			this.cdsConciliacion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsConciliacion.SchemaDef = null;
			this.cdsConciliacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsConciliacion_BeforeFill);
			// 
			// btSubirR
			// 
			this.btSubirR.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSubirR.Enabled = false;
			this.btSubirR.Location = new System.Drawing.Point(248, 32);
			this.btSubirR.Name = "btSubirR";
			this.btSubirR.Size = new System.Drawing.Size(96, 24);
			this.btSubirR.TabIndex = 28;
			this.btSubirR.Text = "Subir Retencion";
			this.btSubirR.Click += new System.EventHandler(this.btSubirR_Click);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cdvCompra
			// 
			this.cdvCompra.BindingContextCtrl = this;
			this.cdvCompra.DataSet = this.cdsCompra;
			this.cdvCompra.Sort = "Numero";
			this.cdvCompra.TableName = "";
			this.cdvCompra.TableViewName = "Compra";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.Location = new System.Drawing.Point(400, 152);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(384, 160);
			this.ultraGrid1.TabIndex = 26;
			// 
			// listBox1
			// 
			this.listBox1.AllowDrop = true;
			this.listBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.listBox1.HorizontalScrollbar = true;
			this.listBox1.Location = new System.Drawing.Point(8, 16);
			this.listBox1.Name = "listBox1";
			this.listBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.listBox1.Size = new System.Drawing.Size(368, 134);
			this.listBox1.TabIndex = 32;
			// 
			// chkModificaExiste
			// 
			this.chkModificaExiste.Location = new System.Drawing.Point(24, 24);
			this.chkModificaExiste.Name = "chkModificaExiste";
			this.chkModificaExiste.Size = new System.Drawing.Size(136, 32);
			this.chkModificaExiste.TabIndex = 34;
			this.chkModificaExiste.Text = "Modificar Existentes";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkModificaExiste);
			this.groupBox1.Controls.Add(this.btSubir);
			this.groupBox1.Controls.Add(this.cmbTipoFactura);
			this.groupBox1.Controls.Add(this.txtMensajeC);
			this.groupBox1.Location = new System.Drawing.Point(8, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(384, 128);
			this.groupBox1.TabIndex = 36;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Subir Compra de Excel Linea CodCliente es la separacion con el Detalle";
			// 
			// cmbTipoFactura
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance1;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoFactura.DisplayMember = "";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(168, 32);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(128, 21);
			this.cmbTipoFactura.TabIndex = 109;
			this.cmbTipoFactura.ValueMember = "";
			// 
			// txtMensajeC
			// 
			this.txtMensajeC.AcceptsReturn = true;
			this.txtMensajeC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMensajeC.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtMensajeC.Location = new System.Drawing.Point(8, 64);
			this.txtMensajeC.Multiline = true;
			this.txtMensajeC.Name = "txtMensajeC";
			this.txtMensajeC.Size = new System.Drawing.Size(368, 56);
			this.txtMensajeC.TabIndex = 30;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.btLeer);
			this.groupBox2.Controls.Add(this.txtMensajeR);
			this.groupBox2.Controls.Add(this.btSubirR);
			this.groupBox2.Location = new System.Drawing.Point(400, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(384, 128);
			this.groupBox2.TabIndex = 37;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Subir Rretenciones Execl linea";
			// 
			// btLeer
			// 
			this.btLeer.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btLeer.Location = new System.Drawing.Point(24, 32);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(96, 24);
			this.btLeer.TabIndex = 30;
			this.btLeer.Text = "Leer Archivo";
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// txtMensajeR
			// 
			this.txtMensajeR.AcceptsReturn = true;
			this.txtMensajeR.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMensajeR.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtMensajeR.Location = new System.Drawing.Point(8, 64);
			this.txtMensajeR.Multiline = true;
			this.txtMensajeR.Name = "txtMensajeR";
			this.txtMensajeR.Size = new System.Drawing.Size(368, 56);
			this.txtMensajeR.TabIndex = 29;
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox3.Controls.Add(this.listBox1);
			this.groupBox3.Location = new System.Drawing.Point(8, 152);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(384, 160);
			this.groupBox3.TabIndex = 38;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Mensajes de Subida de Archivos";
			// 
			// P_ImportaCompraLinea
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(792, 317);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "P_ImportaCompraLinea";
			this.Text = "Importa Compra y  Retencion por Linea";
			this.Load += new System.EventHandler(this.P_ImportaCompraLinea_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsConciliacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensajeC)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtMensajeR)).EndInit();
			this.groupBox3.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion		
		#region Botones y Funciones del sistema
		int control=0;
		
		private void P_ImportaCompraLinea_Load(object sender, System.EventArgs e)
		{
			string stCombo = "Select idTipoFactura,Nombre from compranumero order by idtipofactura";
			cmbTipoFactura.DataSource = Funcion.dvProcedimiento(cdsConciliacion,stCombo);		
			cmbTipoFactura.DisplayMember = "Nombre";
			cmbTipoFactura.ValueMember = "idTipoFactura";
			cmbTipoFactura.Rows.Band.Columns["idTipoFactura"].Hidden = true;
		}
		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsConciliacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsConciliacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		#endregion Botones y Funciones del sistema	
		string vObra = "";string vCap = "";string vSCap = "";
		private void btSubir_Click(object sender, System.EventArgs e)
		{
			this.txtMensajeC.Text = "Formato: \nNumero, Fecha, Ruc, ContadoCredito,....CodArticulo, Cantidad, Codigo, Precio";
			#region Verifica TipoFactura
			this.errorProvider.SetError(this.cmbTipoFactura, "");
			listBox1.Items.Clear();
			if (cmbTipoFactura.Value == null || cmbTipoFactura.Value == System.DBNull.Value || (int) cmbTipoFactura.Value <= 0)
			{								
				this.errorProvider.SetError(this.cmbTipoFactura, "Ingrese Articulo");
				return;
			}
			#endregion Verifica TipoFactura			
			#region Variables de Inicio			
			int IdCompra = 0;
			string stSelect ="";
			string stInsert="";
			string stInsertDC="";	
			string stUpdate ="";
			int idAux = 0;
			int idCliente=0;


			C1.Data.C1DataRow drCompra = this.cdsCompra.TableViews["Compra"].AddNew();
			C1.Data.C1DataRow drDetCompra = this.cdsCompra.TableViews["DetCompra"].AddNew();
			#endregion Variables de Inicio
			#region Mensaje Confirmacion
			if (MessageBox.Show("¿Desea Subir Compras de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
			{
				listBox1.Items.Add("Subida Cancelada por Usuario");
				return;	
			}
			#endregion Mensaje Confirmacion
			#region Apertura de Archivo de Excel
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK != this.openFileDialog.ShowDialog())
				return;
			this.stArchivo = this.openFileDialog.FileName.ToString();
			this.stArchivo = this.stArchivo.ToUpper();
			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(this.stArchivo);
			}
			catch 
			{
				//MessageBox.Show(exception1.Message);
				//listBox1.Items.Add("Abrir Archivo : "+exception1.Message);
				return;
			}
			#endregion				
			#region Verificacion de Titulos
			string[] stCampo = new string[200];
			XLSheet sheet1 = book1.Sheets[0];
			XLCell cell1 = sheet1.GetCell(0, 0);
			int iNumColumnas = 0;
			control=0;
			for (int i = 0; i < sheet1.Columns.Count; i++)
			{
				cell1 = sheet1.GetCell(0, i);				
				if (cell1 != null && cell1.Value != null)
				{
					try
					{
						stCampo[i] = (string) cell1.Value;
					}
					catch (Exception exception2)
					{
						MessageBox.Show(exception2.Message, "Error al leer titulo");
						//listBox1.Items.Add("Error al leer titulo: "+exception2.Message);
						return;
					}
					if (i == 0)
						if (stCampo[i].ToUpper().Trim() != "NUMERO")
						{
							listBox1.Items.Add("Primera Columna debe Ser Numero");
							return;
						}
					if (i == 1)
						if (stCampo[i].ToUpper().Trim() != "RUC")
						{
							listBox1.Items.Add("Segunda Columna debe Ser RUC");
							return;
						}
					stCampo[i] = TitulosExcel(stCampo[i].ToUpper().Trim());
					if(stCampo[i].ToUpper().Trim() == "IDARTICULO")control = iNumColumnas;
					try
					{
						if (control==0)
							drCompra[stCampo[i]].GetType();
						else
							drDetCompra[stCampo[i]].GetType();
						iNumColumnas++;
					}
					catch
					{
						listBox1.Items.Add("No existe campo "+ stCampo[i].Trim()	+ ".\nPrimera fila debe tener los campos de Compras a importar.");
						return;
					}					
				}
				else
				{
					listBox1.Items.Add("No debe haber columnas sin titulos.\nPrimera fila debe tener los campos de Compras a importar.");
					return;
				}					
			}
			#endregion Verificacion de Titulos			
			#region Lectura de Filas
			for (int iFila = 1; iFila < sheet1.Rows.Count; iFila++)
			{
				 vObra = "0";				 vCap = "0";				 vSCap = "0";
				cell1 = sheet1.GetCell(iFila, 0);					
				if ((cell1 == null) || (cell1.Value == null)) continue;
				#region Verifica si existe compra
				XLCell cell2 = sheet1.GetCell(iFila, 1);					
				stSelect = string.Format("Select count(c.idcompra) from compra c inner join cliente cl on cl.idcliente = c.idcliente "
					+"where idtipofactura = '{0}' and numero = '{1}' and cl.ruc = '{2}' and c.borrar = 0",cmbTipoFactura.Value, cell1.Value,cell2.Value);
				int cuenta = Funcion.iEscalarSQL(cdsConciliacion,stSelect);
				if (cuenta > 1 )
				{
					listBox1.Items.Add(string.Format("Existen Varias registros con ese numero: '{0}' y ruc: '{1}'",cell1.Value,cell2.Value));
					continue;								
				}
				#endregion Verifica si existe compra				
				#region Actualiza la Compra
				if (cuenta == 1)
				{
					string stSelect1=  string.Format("Select c.idcompra from compra c inner join cliente cl on cl.idcliente = c.idcliente "
						+"where idtipofactura = '{0}' and numero = '{1}' and cl.ruc = '{2}' and c.borrar = 0",cmbTipoFactura.Value,cell1.Value,cell2.Value);
					IdCompra = Funcion.iEscalarSQL(cdsConciliacion,stSelect1);
					if (chkModificaExiste.Checked == true)
					{	
						//campo para verificar si se trata de la misma compra para borrar el detalle
						if (idAux != IdCompra)
						{
							string stBorraDetalle="Delete detcompra where idcompra = "+IdCompra;
							Funcion.EjecutaSQL(cdsConciliacion,stBorraDetalle);
							idAux = IdCompra;
						}
						#region Crea Titulos y pone valores del Update compra
						stUpdate = "Update Compra set ";
						for (int i = 0; i < control; i++)
						{
							XLCell cell3 = sheet1.GetCell(0, i);	
							cell1 = sheet1.GetCell(iFila, i);	
							if (cell3.Value.ToString().ToUpper().Trim() == "RUC" ||cell3.Value.ToString().ToUpper().Trim() == "IDCLIENTE")
							{
								cell3.Value="idCliente";
								#region Captura el idCliente segun el ruc								
								string stCliente = string.Format("select isnull(count(idcliente),0) from cliente where ruc = '{0}'",cell1.Value);
								int Cuenta = Funcion.iEscalarSQL(cdsConciliacion,stCliente);
								if (Cuenta >= 1)
								{
									if (Cuenta > 1)
										listBox1.Items.Add("Existen varios Registros con ruc, Se inserta el primero: "+cell1.Value+" En linea de excel "+(iFila+1));										
									stCliente = string.Format("select isnull(idcliente,0) from cliente where ruc = '{0}'",cell1.Value);
									int idCli = Funcion.iEscalarSQL(cdsConciliacion,stCliente);
									cell1.Value= idCli;
									idCliente=idCli;
								}
								if (Cuenta <= 0 )
								{
									listBox1.Items.Add("No existe Registro con ruc, No se Actualiza el registro: "+cell1.Value+" En linea de excel "+(iFila+1));
									cell1.Value = 0;
									idCliente=0;
								}								
								#endregion Captura el idCliente segun el ruc
							}
							if(control - 1 == i)
								stUpdate +=" "+cell3.Value +" = '"+cell1.Value+"' where idcompra = "+IdCompra; 
							else
								stUpdate +=" "+cell3.Value +" = '"+cell1.Value+"' ,"; 
						}	
						#endregion Crea Titulos del Update compra		
						if (idCliente == 0)continue;
						Funcion.EjecutaSQL(cdsConciliacion,stUpdate);
						//listBox1.Items.Add("UPDATE "+stUpdate);
					}
				}				
				#endregion Actualiza la Compra
				#region Ingresa Nueva Compra 				
				if (cuenta == 0) 
				{	
					#region Crea Cabecera Compra
					#region Crea Titulos del Insert compra
					stInsert="Insert Into Compra(";
					for (int i = 0; i < control; i++)
					{
						XLCell cell3 = sheet1.GetCell(0, i);	
						if (cell3.Value.ToString().ToUpper().Trim() == "RUC")
							cell3.Value="idCliente";
						if(control - 1 == i)
							stInsert +=" "+cell3.Value +",idTipoFactura )"; 
						else
							stInsert +=" "+cell3.Value +" ,"; 
					}	
					#endregion Crea Titulos del Insert compra		
					#region Ingreso Valores Insert Compra
					stInsert +=" Values ( ";
					for (int iCol = 0; iCol < control; iCol++)
					{							
						cell1 = sheet1.GetCell(iFila, iCol);
						if ((cell1 == null) || (cell1.Value == null)) continue;
						try
						{	
							#region Captura el idCliente segun el ruc
							if (iCol == 1)
							{
								string stCliente = string.Format("select isnull(count(idcliente),0) from cliente where ruc = '{0}'",cell1.Value);
								int Cuenta = Funcion.iEscalarSQL(cdsConciliacion,stCliente);
								if (Cuenta >= 1)
								{
									if (Cuenta > 1)
										listBox1.Items.Add("Existen varios Registros con ruc, Se inserta el primero: "+cell1.Value+" En linea de excel "+(iFila+1));										
									stCliente = string.Format("select top 1 isnull(idcliente,0) from cliente where ruc = '{0}'",cell1.Value);
									int idCli = Funcion.iEscalarSQL(cdsConciliacion,stCliente);
									cell1.Value= idCli;
									idCliente = idCli;
								}									
								if (Cuenta <= 0)
								{
									listBox1.Items.Add("No existe Registro con ruc, No se Inserta registro: "+cell1.Value+" En linea de excel "+(iFila+1));
									cell1.Value = 0;
									idCliente=0;
								}
							}
							#endregion Captura el idCliente segun el ruc
							if(control - 1 == iCol)
								stInsert +=" '"+cell1.Value +"','"+cmbTipoFactura.Value+"' ) "+ 
									"Select Top 1 idCompra From Compra Order By idCompra DESC"; 
							else
								stInsert +=" '"+cell1.Value +"' ,"; 						
						}
						catch(Exception ex)
						{
							listBox1.Items.Add("Error Verifica Cliente: "+ex.Message+" En linea de excel "+(iFila+1));
							return;
						}
					}
					#endregion Ingreso Valores 				
					#region Inserta nueva compra a base sql
					try
					{
						if (idCliente == 0)continue;
						IdCompra =  Funcion.iEscalarSQL(cdsConciliacion,stInsert);	
						idAux = IdCompra;			
					}
					catch(Exception ex)
					{
						listBox1.Items.Add("No se pudo insertar nuevo registro: "+ex.Message+" En linea de excel "+(iFila+1));
						return;
					}
					#endregion Inserta nueva compra a base sql
					#endregion Crea Cabecera Compra	
				}														
				#endregion Ingresa Nueva Compra
				#region Crea Detalle de compra
				if (IdCompra > 0)
				{
					#region Crea Titulos del Insert DetCompra
					stInsertDC="Insert Into DetCompra(";
					for (int i = control; i < sheet1.Columns.Count; i++)
					{
						XLCell cell3 = sheet1.GetCell(0, i);					
						cell3.Value = TitulosExcel(cell3.Value.ToString().ToUpper().Trim());
						if(sheet1.Columns.Count - 1 == i)
							stInsertDC +=" "+cell3.Value +",idcompra )"; 
						else
							stInsertDC +=" "+cell3.Value +" ,"; 
					}	
					#endregion Crea Titulos del Insert DetCompra		
					#region Ingreso Valores Insert DetCompra
					stInsertDC +=" Values ( ";
					for (int iCol = control; iCol < sheet1.Columns.Count; iCol++)
					{					
						cell1 = sheet1.GetCell(iFila, iCol);

						if ((cell1 == null) || (cell1.Value == null)) continue;
						try
						{						
							#region Captura el idArticulo segun el CodArticulo
							XLCell cell5 = sheet1.GetCell(0, iCol);

							//llena campos
							cell1.Value = gValores(cell5.Value.ToString().Trim().ToUpper(),iFila,cell1.Value.ToString());

							#endregion Captura el idArticulo segun el CodArticulo
							if(sheet1.Columns.Count - 1 == iCol)
								stInsertDC +=" '"+cell1.Value +"', '"+IdCompra+"' )"; 
							else
								stInsertDC +=" '"+cell1.Value +"' ,"; 						
						}
						catch(Exception ex)
						{
							listBox1.Items.Add("Error Verifica Articulo: "+ex.Message+" En linea de excel "+(iFila+1));
							return;
						}
					}
					try
					{		
						string activaTrigger="ALTER TABLE DETCOMPRA DISABLE TRIGGER SALDOCOSTOI "+
                                 "ALTER TABLE DETCOMPRA DISABLE TRIGGER SALDOCOSTO";						
						Funcion.EjecutaSQL(cdsConciliacion,activaTrigger);
						Funcion.EjecutaSQL(cdsConciliacion,stInsertDC);
										activaTrigger="ALTER TABLE DETCOMPRA enable TRIGGER SALDOCOSTOI "+
										"ALTER TABLE DETCOMPRA enable TRIGGER SALDOCOSTO";
						Funcion.EjecutaSQL(cdsConciliacion,activaTrigger);

					}
					catch(Exception ex)
					{
						listBox1.Items.Add("No se pudo insertar nuevo registro Detalle Compra: "+ex.Message);
						return;
					}
					#endregion Ingreso Valores Insert DetCompra		
				}
				#endregion Crea Detalle de compra
				Funcion.EjecutaSQL(cdsConciliacion,"Exec facturaTotal "+IdCompra);
			}
			#endregion Lectura de Filas		
			MessageBox.Show("Factura Importada", "Información");
		}
		private void btSubirR_Click(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Subir Retenciones de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No)
			{
				btSubirR.Enabled=false;
				return;
			}
			subir();
			btSubirR.Enabled=false;
		}

	
		#region Funciones Personalizadas
		
		//Declaracion de Variables Globales
		string stArchivo = "";
		DataSet miDataSet = new DataSet();

		private int AbrirArchivo()
		{
			int abre=0; 
				this.openFileDialog.Filter = "xls (*.xls)|*.xls";						
			this.openFileDialog.InitialDirectory = Application.ExecutablePath;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				stArchivo = this.openFileDialog.FileName.ToString();
				btSubirR.Enabled=true;
			}
			else
			{
				listBox1.Items.Add("Cancelado por el usuario");				
				this.Cursor = Cursors.Default;
				btSubirR.Enabled=false;
				abre=1;
			}
			return abre;
		}
		
		private void LeerDatosExcel()
		{
			//boton de facturas para subir del excel a la grilla
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + stArchivo + ";" +
				"Extended Properties='Excel 8.0;'";		
			string stHoja = "select * from [Hoja1$]";
			OleDbDataAdapter myData = new OleDbDataAdapter(stHoja, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener cerrado el archivo " + stArchivo + " en Excel",
					ex.Message);
				btSubirR.Enabled=false;
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}
		
		private void subir()
		{
			string stExec ="";
			Cursor = Cursors.WaitCursor;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{					
				try
				{	
					string ruc = dr.Cells[0].Value.ToString();
					string numCompra = dr.Cells[1].Value.ToString(); 
					float ivaretenido = float.Parse (dr.Cells[2].Value.ToString());
					float ivaretenido1 = float.Parse (dr.Cells[3].Value.ToString());
					string CodIva = dr.Cells[4].Value.ToString(); 
					string CodIva1 = dr.Cells[5].Value.ToString();  
					float facturaIva12 = float.Parse (dr.Cells[6].Value.ToString());
					float facturaIva121 = float.Parse (dr.Cells[7].Value.ToString()); 
					string AutorizaRet = dr.Cells[8].Value.ToString();  
					string Serie = dr.Cells[9].Value.ToString(); 
					string codigo = dr.Cells[10].Value.ToString();
					DateTime dtDesde = DateTime.Parse(dr.Cells[11].Value.ToString());					
					float Base = float.Parse (dr.Cells[12].Value.ToString()); 
					string codRetDetalle = dr.Cells[13].Value.ToString(); 
				

				stExec = string.Format("exec P_PresupuestoImportaRet '{0}','{1}','{2}','{3}','{4}','{5}','{6}',"+
					"'{7}','{8}','{9}','{10}','{11}','{12}','{13}'",
				ruc , numCompra , ivaretenido, ivaretenido1 ,CodIva , CodIva1 , facturaIva12 ,
				facturaIva121, AutorizaRet , Serie , codigo,dtDesde.ToString("yyyyMMdd") ,Base , codRetDetalle );
				}
				catch(Exception ex)
				{
					listBox1.Items.Add("Error al leer Columnas de excel.."+ex);
					return;
				}

				try
				{																		
					string msj = Funcion.sEscalarSQL(cdsConciliacion,stExec); 
					if(msj.Length!=0)
						listBox1.Items.Add(" "+msj);
				}
				catch(Exception ex)
				{
					MessageBox.Show("Error al importar presupuesto...." + ex.Message);
				}
			}
			Cursor = Cursors.Default;
			MessageBox.Show("Proceso Terminado...");
		}
		

		private string TitulosExcel(string campo)
		{			
			string retorna=campo;
			if (campo == "CLIENTE")
			{
				retorna = "idCliente";
			}
			if (campo == "RUC")
			{
				retorna = "idCliente";
			}
			if (campo == "CODARTICULO")
			{
				retorna =  "idArticulo";				
			}
			if (campo == "OBRA")
			{
				retorna =  "idPresupuestoC";				
			}
			if (campo == "CAPITULO")
			{
				retorna =  "idCapitulo";				
			}
			if (campo == "SUBCAPITULO")
			{
				retorna =  "idSubCapitulo";				
			}
			if (campo == "RUBRO")
			{
				retorna =  "idPresupuestoCDet";				
			}

			return retorna;
		}
		private string gValores(string campo, int fila, string valor)
		{
			string retorna=valor;
			string stArt="";
			#region Asigna id Arrticulo
			if (campo == "IDARTICULO")
			{
				int idArt = 0;
				stArt = string.Format("select isnull(count(idArticulo),0) from Articulo where codigo = '{0}'",valor);
				int Cuenta = Funcion.iEscalarSQL(cdsConciliacion,stArt);
				if (Cuenta == 1)
				{
					stArt = string.Format("select isnull(idArticulo,0) from Articulo where codigo = '{0}'",valor);
					idArt = Funcion.iEscalarSQL(cdsConciliacion,stArt);
					retorna = idArt.ToString();
				}
				else
				{
					listBox1.Items.Add("No existe Articulo con Codigo: "+valor+" En linea de excel "+(fila+1));									
					retorna = "0";
				}				
			}
			#endregion Asigna id Arrticulo
			#region Asigna id Obra
			//obra
			if (campo == "OBRA" || campo == "IDPRESUPUESTOC")
			{
				int idArt = 0;
				stArt = string.Format("select isnull(count(idPresupuestoC),0) from P_Presupuesto where Obra = '{0}'",valor);
				int Cuenta = Funcion.iEscalarSQL(cdsConciliacion,stArt);
				if (Cuenta == 1)
				{
					stArt = string.Format("select isnull(idPresupuestoC,0) from P_Presupuesto where Obra = '{0}'",valor);
					idArt = Funcion.iEscalarSQL(cdsConciliacion,stArt);
					retorna = idArt.ToString();
					vObra = retorna;
				}
				else
				{
					listBox1.Items.Add("No existe Obras con Codigo: "+valor+" En linea de excel "+(fila+1));									
					retorna = "0";
				}				
			}
			#endregion Asigna id Obra
			#region Asigna id capitulo
			//capitulo
			if (campo == "IDCAPITULO")
			{
				int idArt = 0;
				stArt = string.Format("select isnull(count(idCapitulo),0) from P_PresupuestoCapitulo where Capitulo = '{0}' and idPresupuestoC = '{1}'",valor,vObra);
				int Cuenta = Funcion.iEscalarSQL(cdsConciliacion,stArt);
				if (Cuenta == 1)
				{
					stArt = string.Format("select isnull(idCapitulo,0) from P_PresupuestoCapitulo where Capitulo = '{0}' and idPresupuestoC = '{1}'",valor,vObra);
					idArt = Funcion.iEscalarSQL(cdsConciliacion,stArt);
					retorna = idArt.ToString();
					vCap=retorna;
				}
				else
				{
					listBox1.Items.Add("No existe Capitulos con Codigo: "+valor+" En linea de excel "+(fila+1));									
					retorna = "0";
				}				
			}
			#endregion Asigna id capitulo
			#region Asigna id SUBCAPITULO
			//SUBCAPITULO
			if (campo == "IDSUBCAPITULO")
			{
				int idArt = 0;
				stArt = string.Format("select isnull(count(idSubCapitulo),0) from P_PresupuestosubCapitulo where SubCapitulo = '{0}' and idCapitulo = '{1}'",valor,vCap);
				int Cuenta = Funcion.iEscalarSQL(cdsConciliacion,stArt);
				if (Cuenta == 1)
				{
					stArt = string.Format("select isnull(idSubCapitulo,0) from P_PresupuestosubCapitulo where SubCapitulo = '{0}' and idCapitulo = '{1}'",valor,vCap);
					idArt = Funcion.iEscalarSQL(cdsConciliacion,stArt);
					retorna = idArt.ToString();
					vSCap=retorna;
				}
				else
				{
					listBox1.Items.Add("No existe SubCapitulo con Codigo: "+valor+" En linea de excel "+(fila+1));									
					retorna = "0";
				}				
			}
			#endregion Asigna id SUBCAPITULO
			#region Asigna id RUBRO
			//RUBRO
			if (campo == "RUBRO" || campo == "IDPRESUPUESTOCDET")
			{
				int idArt = 0;
				stArt = string.Format("select isnull(count(idPresupuestoCDet),0) from P_PresupuestoDetalle where Rubro = '{0}' and idSubCapitulo = '{1}'",valor,vSCap);
				int Cuenta = Funcion.iEscalarSQL(cdsConciliacion,stArt);
				if (Cuenta == 1)
				{
					stArt = string.Format("select isnull(idPresupuestoCDet,0) from P_PresupuestoDetalle where Rubro = '{0}' and idSubCapitulo = '{1}'",valor,vSCap);
					idArt = Funcion.iEscalarSQL(cdsConciliacion,stArt);
					retorna = idArt.ToString();
				}
				else
				{
					listBox1.Items.Add("No existe Rubros con Codigo: "+valor+" En linea de excel "+(fila+1));									
					retorna = "0";
				}				
			}
			#endregion Asigna id RUBRO
			return retorna;
		}
		
		#endregion Funciones Personalizadas			

		private void btLeer_Click(object sender, System.EventArgs e)
		{
			listBox1.Items.Clear();
			string texto ="NoRuc,	NoFactura, ValorRetIva, ValorRetIva1, CodigoRetIva, CodigoRetIva1, SubTotal, SubTotalIva1, Autorizacion, Serie, NoRetencion, Fecha, BaseRenta, CodRetencion";
			this.txtMensajeR.Text = "Formato: \n "+texto;
			if (AbrirArchivo()!= 0)return;			 
			LeerDatosExcel();							
		}

		
	}
}
