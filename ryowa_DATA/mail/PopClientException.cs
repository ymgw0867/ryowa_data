using System;
using System.Windows.Forms;

namespace ryowa_DATA.mail
{
	/// <summary>
	/// PopClient の例外クラスです。
	/// </summary>
	public class PopClientException : Exception
	{
		/// <summary>
		/// コンストラクタです。
		/// </summary>
		public PopClientException()
		{
		}

		/// <summary>
		/// コンストラクタです。
		/// </summary>
		/// <param name="message"></param>
		public PopClientException(string message) : base(message)
		{
            MessageBox.Show(message);
		}

		/// <summary>
		/// コンストラクタです。
		/// </summary>
		/// <param name="message"></param>
		/// <param name="innerException"></param>
		public PopClientException(string message, Exception innerException) : base(message, innerException)
		{
		}
	}
}
