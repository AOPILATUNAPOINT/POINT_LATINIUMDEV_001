using System;
using System.Windows.Forms;
using System.Data;
using System.Drawing;

namespace Latinium.Services.FamilyFriends
{
	public class DataGridNumericTextBoxColumn : DataGridTextBoxColumn
	{
		private TextBox _textBox;

		protected override void Edit(
			CurrencyManager source,
			int rowNum,
			Rectangle bounds,
			bool readOnly,
			string instantText,
			bool cellIsVisible)
		{
			base.Edit(source, rowNum, bounds, readOnly, instantText, cellIsVisible);

			_textBox = this.TextBox;

			if (_textBox != null)
			{
				_textBox.KeyPress -= new KeyPressEventHandler(OnKeyPress);
				_textBox.KeyPress += new KeyPressEventHandler(OnKeyPress);
			}
		}

		private void OnKeyPress(object sender, KeyPressEventArgs e)
		{
			// permitir números
			if (char.IsDigit(e.KeyChar))
				return;

			// permitir backspace
			if (e.KeyChar == (char)8)
				return;

			// permitir punto decimal
			if (e.KeyChar == '.')
			{
				TextBox tb = (TextBox)sender;

				// evitar múltiples puntos
				if (tb.Text.IndexOf('.') == -1)
					return;
			}

			// bloquear todo lo demás
			e.Handled = true;
		}
	}
}
