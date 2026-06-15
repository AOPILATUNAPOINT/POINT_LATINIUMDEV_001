using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Latinium.Models;
using Latinium.Services;
using Latinium.Loaders;

namespace Latinium
{
	public class frmRequerimientoDepartamentoGrupo : Form
	{
		private ComboBox cmbDepartamentos;
		private Label lblMensaje;
		private DataGrid dgAsociaciones;
		private Label label1;
		private CheckBox chkMaestro;
		private Button btnGuardar;

		private string connectionString = frmRequerimiento.sconexionPoint;
		private int idDepartamentoActual = -1;

		public frmRequerimientoDepartamentoGrupo()
		{
			InitializeComponent();
			this.Load += new EventHandler(this.frmRequerimientoDepartamentoGrupo_Load);
			this.Resize += new EventHandler(this.frmRequerimientoDepartamentoGrupo_Resize);
		}

		private void InitializeComponent()
		{
			this.cmbDepartamentos = new System.Windows.Forms.ComboBox();
			this.lblMensaje = new System.Windows.Forms.Label();
			this.dgAsociaciones = new System.Windows.Forms.DataGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dgAsociaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDepartamentos
			// 
			this.cmbDepartamentos.Location = new System.Drawing.Point(144, 16);
			this.cmbDepartamentos.Name = "cmbDepartamentos";
			this.cmbDepartamentos.Size = new System.Drawing.Size(424, 24);
			this.cmbDepartamentos.TabIndex = 0;
			this.cmbDepartamentos.SelectedIndexChanged += new System.EventHandler(this.cmbDepartamentos_SelectedIndexChanged);
			// 
			// lblMensaje
			// 
			this.lblMensaje.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lblMensaje.Location = new System.Drawing.Point(20, 50);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(400, 21);
			this.lblMensaje.TabIndex = 1;
			// 
			// dgAsociaciones
			// 
			this.dgAsociaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dgAsociaciones.DataMember = "";
			this.dgAsociaciones.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgAsociaciones.Location = new System.Drawing.Point(20, 80);
			this.dgAsociaciones.Name = "dgAsociaciones";
			this.dgAsociaciones.Size = new System.Drawing.Size(548, 300);
			this.dgAsociaciones.TabIndex = 2;
			this.dgAsociaciones.CurrentCellChanged += new System.EventHandler(this.dgAsociaciones_CurrentCellChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(20, 16);
			this.label1.Name = "label1";
			this.label1.TabIndex = 3;
			this.label1.Text = "Ticketera:";
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Location = new System.Drawing.Point(448, 390);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(120, 30);
			this.btnGuardar.TabIndex = 4;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// frmRequerimientoDepartamentoGrupo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(592, 430);
			this.Controls.Add(this.cmbDepartamentos);
			this.Controls.Add(this.lblMensaje);
			this.Controls.Add(this.dgAsociaciones);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnGuardar);
			this.Name = "frmRequerimientoDepartamentoGrupo";
			this.Text = "Requerimiento Departamento Grupo";
			this.Load += new System.EventHandler(this.frmRequerimientoDepartamentoGrupo_Load_1);
			((System.ComponentModel.ISupportInitialize)(this.dgAsociaciones)).EndInit();
			this.ResumeLayout(false);

		}

		private void frmRequerimientoDepartamentoGrupo_Load(object sender, EventArgs e)
		{
			DepartamentoLoader.CargarDepartamentos(this.cmbDepartamentos, this.connectionString, this.lblMensaje);
			ConfigurarDataGrid();
		}

		private void cmbDepartamentos_SelectedIndexChanged(object sender, EventArgs e)
		{
			RequerimientoDepartamento departamento = this.cmbDepartamentos.SelectedItem as RequerimientoDepartamento;

			if (departamento != null && departamento.IdRequerimientoDepartamento > 0)
			{
								this.idDepartamentoActual = departamento.IdRequerimientoDepartamento;

				AsociacionesLoader.CargarAsociaciones(
					departamento.IdRequerimientoDepartamento,
					this.dgAsociaciones,
					this.connectionString,
					this.lblMensaje,
					new AsociacionesLoader.MostrarCheckMaestroDelegate(this.MostrarOcultarCheckMaestro));

				this.lblMensaje.Text = "Asociaciones cargados para el departamento seleccionado.";


			}
			else
			{
				this.dgAsociaciones.DataSource = null;
				this.OcultarCheckMaestro();
				this.lblMensaje.Text = "Seleccione un departamento.";

			}
			ReposicionarChkMaestro();
		}

		private void MostrarOcultarCheckMaestro(bool mostrar)
		{
			if (mostrar)
			{
				if (this.chkMaestro == null)
				{
					this.chkMaestro = new CheckBox();
					this.chkMaestro.Text = "Seleccionar/Deseleccionar todos";
					this.chkMaestro.Size = new Size(180, 18);
					this.chkMaestro.Checked = false;
					this.chkMaestro.Cursor = Cursors.Hand;

					this.chkMaestro.CheckedChanged += new EventHandler(this.chkMaestro_CheckedChanged);
					this.Controls.Add(this.chkMaestro);
				}

				this.chkMaestro.Visible = true;
				this.chkMaestro.BringToFront();
				ReposicionarChkMaestro();
			}
			else
			{
				OcultarCheckMaestro();
			}
		}

		private void OcultarCheckMaestro()
		{
			if (this.chkMaestro != null)
			{
				this.chkMaestro.Visible = false;
				this.chkMaestro.Checked = false;
			}
		}

		private void chkMaestro_CheckedChanged(object sender, EventArgs e)
		{
			ArrayList lista = this.dgAsociaciones.DataSource as ArrayList;
			if (lista != null)
			{
				bool nuevoValor = this.chkMaestro.Checked;

				for (int i = 0; i < lista.Count; i++)
				{
					object item = lista[i];
					System.Reflection.PropertyInfo prop = item.GetType().GetProperty("Estado");
					if (prop != null && prop.CanWrite)
					{
						prop.SetValue(item, nuevoValor, null);
					}
				}

				this.dgAsociaciones.Refresh();
			}
		}

		private void frmRequerimientoDepartamentoGrupo_Resize(object sender, EventArgs e)
		{
			ReposicionarChkMaestro();

			// Ajustar posición del botón Guardar si quieres que se mantenga en la parte inferior derecha
			if (btnGuardar != null)
			{
				btnGuardar.Location = new Point(this.ClientSize.Width - btnGuardar.Width - 20, this.ClientSize.Height - btnGuardar.Height - 20);
			}
		}

		private void ReposicionarChkMaestro()
		{
			if (this.chkMaestro != null && this.chkMaestro.Visible)
			{
				int margen = 10;
				int x = this.dgAsociaciones.Right - this.chkMaestro.Width - margen;
				int y = this.dgAsociaciones.Top - this.chkMaestro.Height - 5;
				this.chkMaestro.Location = new Point(x, y);
			}
		}

		private void ConfigurarDataGrid()
		{
			DataGridTableStyle tableStyle = new DataGridTableStyle();
			tableStyle.MappingName = "ArrayList"; // porque estás usando ArrayList

			DataGridTextBoxColumn colId = new DataGridTextBoxColumn();
			colId.MappingName = "IdGrupo";
			colId.HeaderText = "ID Grupo";
			colId.Width = 100;
			tableStyle.GridColumnStyles.Add(colId);

			DataGridTextBoxColumn colNombre = new DataGridTextBoxColumn();
			colNombre.MappingName = "NombreGrupo";
			colNombre.HeaderText = "Nombre Grupo";
			colNombre.Width = 200;
			tableStyle.GridColumnStyles.Add(colNombre);

			DataGridBoolColumn colEstado = new DataGridBoolColumn();
			colEstado.MappingName = "Estado";
			colEstado.HeaderText = "Estado";
			colEstado.Width = 75;
			tableStyle.GridColumnStyles.Add(colEstado);

			dgAsociaciones.TableStyles.Clear();
			dgAsociaciones.TableStyles.Add(tableStyle);
		}

		private void btnGuardar_Click(object sender, EventArgs e)
		{
			if (this.idDepartamentoActual <= 0)
			{
				MessageBox.Show("Seleccione un departamento válido antes de guardar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			if (this.dgAsociaciones.DataSource == null)
			{
				MessageBox.Show("No hay datos para guardar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			ArrayList listaRoles = new ArrayList();

			ArrayList dataSource = this.dgAsociaciones.DataSource as ArrayList;
			if (dataSource != null)
			{
				for (int i = 0; i < dataSource.Count; i++)
				{
					RequerimientoDepartamentoRol rol = dataSource[i] as RequerimientoDepartamentoRol;
					if (rol != null)
					{
						// Crear una copia con el IdDepartamento actual y el estado posiblemente modificado
						RequerimientoDepartamentoRol rolActualizar = new RequerimientoDepartamentoRol();
						rolActualizar.IdRequerimientoDepartamento = this.idDepartamentoActual;
						rolActualizar.IdGrupo = rol.IdGrupo;
						rolActualizar.NombreGrupo = rol.NombreGrupo;
						rolActualizar.Estado = rol.Estado;

						listaRoles.Add(rolActualizar);
					}
				}

				try
				{
					RequerimientoDepartamentoRolService service = new RequerimientoDepartamentoRolService(connectionString);
					service.GuardarRoles(listaRoles);
					MessageBox.Show("Cambios guardados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al guardar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void dgAsociaciones_CurrentCellChanged(object sender, EventArgs e)
		{
			if (this.dgAsociaciones.CurrentCell.ColumnNumber == 2) // índice columna Estado (0,1,2)
			{
				CurrencyManager cm = (CurrencyManager)this.BindingContext[this.dgAsociaciones.DataSource];
				ArrayList lista = cm.List as ArrayList;

				if (lista != null)
				{
					int rowIndex = this.dgAsociaciones.CurrentCell.RowNumber;
					RequerimientoDepartamentoRol item = (RequerimientoDepartamentoRol)lista[rowIndex];

					if (item != null)
					{
						// Forzar que Estado no sea null o indeterminado
						// En .NET 1.1 bool no es nullable, pero si por algún motivo tienes un valor inválido, lo normal es:
						if (item.Estado != true && item.Estado != false)
						{
							item.Estado = false;
						}

						this.dgAsociaciones.Refresh();
					}
				}
			}
		}

		private void frmRequerimientoDepartamentoGrupo_Load_1(object sender, System.EventArgs e)
		{
		
		}
	}
}
