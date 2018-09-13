using System;
using System.Windows.Forms;

namespace ryowa_DATA.mail
{
	/// <summary>
	/// PopClient �̗�O�N���X�ł��B
	/// </summary>
	public class PopClientException : Exception
	{
		/// <summary>
		/// �R���X�g���N�^�ł��B
		/// </summary>
		public PopClientException()
		{
		}

		/// <summary>
		/// �R���X�g���N�^�ł��B
		/// </summary>
		/// <param name="message"></param>
		public PopClientException(string message) : base(message)
		{
            MessageBox.Show(message);
		}

		/// <summary>
		/// �R���X�g���N�^�ł��B
		/// </summary>
		/// <param name="message"></param>
		/// <param name="innerException"></param>
		public PopClientException(string message, Exception innerException) : base(message, innerException)
		{
		}
	}
}
