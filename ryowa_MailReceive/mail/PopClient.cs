using System;
using System.Collections;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Windows.Forms;
using ryowa_MailReceive.common;

namespace ryowa_MailReceive.mail
{
	/// <summary>
	/// POP ��胁�[������M����N���X�ł��B
	/// </summary>
	public class PopClient : IDisposable
	{
		/// <summary>TCP �ڑ�</summary>
		private TcpClient tcp = null;

		/// <summary>TCP �ڑ�����̃��[�_�[</summary>
		private StreamReader reader = null;

		/// <summary>
		/// �R���X�g���N�^�ł��BPOP�T�[�o�Ɛڑ����܂��B
		/// </summary>
		/// <param name="hostname">POP�T�[�o�̃z�X�g���B</param>
		/// <param name="port">POP�T�[�o�̃|�[�g�ԍ��i�ʏ��110�j�B</param>
		public PopClient(string hostname, int port)
		{
            try
            {
                // �T�[�o�Ɛڑ�
                this.tcp = new TcpClient(hostname, port);
			    this.reader = new StreamReader(this.tcp.GetStream(), Encoding.ASCII);

			    // �I�[�v�j���O��M
			    string s = ReadLine();
			    if (!s.StartsWith("+OK"))
			    {
                    throw new PopClientException("�ڑ����� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
			    }
            }
            catch (Exception e)
            {
                //MessageBox.Show(hostname + Environment.NewLine + e.Message, "�T�[�o�[�ڑ��G���[", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //Environment.Exit(0);

                //��Q�������̓G���[�ŃX�g�b�v�������Ɏ��̃^�C�}�[�Ōp�������Ƃ���@2011/07/20
                global.Msglog = "ERR : " + hostname + " : " + e.Message;
            }
		}

		/// <summary>
		/// ����������s���܂��B
		/// </summary>
		public void Dispose()
		{
			if (this.reader != null)
			{
				((IDisposable)this.reader).Dispose();
				this.reader = null;
			}
			if (this.tcp != null)
			{
				((IDisposable)this.tcp).Dispose();
				this.tcp = null;
			}
		}

		/// <summary>
		/// POP �T�[�o�Ƀ��O�C�����܂��B
		/// </summary>
		/// <param name="username">���[�U���B</param>
		/// <param name="password">�p�X���[�h�B</param>
		public void Login(string username, string password)
		{
			// ���[�U�����M
			SendLine("USER " + username);
			string s = ReadLine();
			if (!s.StartsWith("+OK"))
			{
				throw new PopClientException("USER ���M���� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
			}

			// �p�X���[�h���M
			SendLine("PASS " + password);
			s = ReadLine();
			if (!s.StartsWith("+OK"))
			{
				throw new PopClientException("PASS ���M���� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
			}
		}

        ///--------------------------------------------------------------------
        /// <summary>
        ///     POP �T�[�o�ɗ��܂��Ă��郁�[�������擾���܂��B
        ///     : 2018/12/15</summary>
        /// <returns>
        ///     ���[���� </returns>
        ///--------------------------------------------------------------------
        public int GetStat()
        {
            // STAT ���M
            SendLine("STAT");
            string s = ReadLine();
            if (!s.StartsWith("+OK"))
            {
                throw new PopClientException("STAT ���M���� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
            }

            // �T�[�o�ɂ��܂��Ă��郁�[���̐����擾
            string[] c = s.Split(' ');
            return Utility.StrtoInt(c[1]);

            //ArrayList list = new ArrayList();
            //while (true)
            //{
            //    s = ReadLine();
            //    if (s == ".")
            //    {
            //        // �I�[�ɓ��B
            //        break;
            //    }
            //    // ���[���ԍ������݂̂����o���i�[
            //    int p = s.IndexOf(' ');
            //    if (p > 0)
            //    {
            //        s = s.Substring(0, p);
            //    }
            //    list.Add(s);
            //}
            //return list;
        }

        /// <summary>
        /// POP �T�[�o�ɗ��܂��Ă��郁�[���̃��X�g���擾���܂��B
        /// </summary>
        /// <returns>System.String ���i�[���� ArrayList�B</returns>
        public ArrayList GetList()
		{
			// LIST ���M
			SendLine("LIST");
			string s = ReadLine();
			if (!s.StartsWith("+OK"))
			{
				throw new PopClientException("LIST ���M���� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
			}

			// �T�[�o�ɂ��܂��Ă��郁�[���̐����擾
			ArrayList list = new ArrayList();
			while (true)
			{
				s = ReadLine();
				if (s == ".")
				{
					// �I�[�ɓ��B
					break;
				}
				// ���[���ԍ������݂̂����o���i�[
				int p = s.IndexOf(' ');
				if (p > 0)
				{
					s = s.Substring(0, p);
				}
				list.Add(s);
			}
			return list;
		}

		/// <summary>
		/// POP �T�[�o���烁�[���� 1�擾���܂��B
		/// </summary>
		/// <param name="num">GetList() ���\�b�h�Ŏ擾�������[���̔ԍ��B</param>
		/// <returns>���[���̖{�́B</returns>
		public string GetMail(string num)
		{
			// RETR ���M
			SendLine("RETR " + num);
			string s = ReadLine();
			if (!s.StartsWith("+OK"))
			{
				throw new PopClientException("RETR ���M���� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
			}

			// ���[���擾
			StringBuilder sb = new StringBuilder();
			while (true)
			{
				s = ReadLine();
				if (s == ".")
				{
					// "." �݂̂̏ꍇ�̓��[���̏I�[��\��
					break;
				}
				sb.Append(s);
				sb.Append("\r\n");
			}
			return sb.ToString();
		}

		/// <summary>
		/// POP �T�[�o�̃��[���� 1�폜���܂��B
		/// </summary>
		/// <param name="num">GetList() ���\�b�h�Ŏ擾�������[���̔ԍ��B</param>
		public void DeleteMail(string num)
		{
			// DELE ���M
			SendLine("DELE " + num);
			string s = ReadLine();
			if (!s.StartsWith("+OK"))
			{
				throw new PopClientException("DELE ���M���� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
			}
		}

		/// <summary>
		/// POP �T�[�o�Ɛؒf���܂��B
		/// </summary>
		public void Close()
		{
			// QUIT ���M
			SendLine("QUIT");
			string s = ReadLine();
			if (!s.StartsWith("+OK"))
			{
				throw new PopClientException("QUIT ���M���� POP �T�[�o�� \"" + s + "\" ��Ԃ��܂����B");
			}

			((IDisposable)this.reader).Dispose();
			this.reader = null;
			((IDisposable)this.tcp).Dispose();
			this.tcp = null;
		}

		/// <summary>
		/// POP �T�[�o�ɃR�}���h�𑗐M���܂��B
		/// </summary>
		/// <param name="s">���M���镶����B</param>
		private void Send(string s)
		{
			global.Msglog = "���M: " + s;
			byte[] b = Encoding.ASCII.GetBytes(s);
			this.tcp.GetStream().Write(b, 0, b.Length);
		}

		/// <summary>
		/// POP �T�[�o�ɃR�}���h�𑗐M���܂��B�����ɉ��s��t�����܂��B
		/// </summary>
		/// <param name="s">���M���镶����B</param>
		private void SendLine(string s)
		{
            global.Msglog = "���M: " + s;
			byte[] b = Encoding.ASCII.GetBytes(s + "\r\n");
			this.tcp.GetStream().Write(b, 0, b.Length);
		}

		/// <summary>
		/// POP �T�[�o���� 1�s�ǂݍ��݂܂��B
		/// </summary>
		/// <returns>�ǂݍ��񂾕�����B</returns>
		private string ReadLine()
		{
			string s = this.reader.ReadLine();
            global.Msglog = "��M: " + s;
			return s;
		}

		/// <summary>
		/// �`�F�b�N�p�ɃR���\�[���ɏo�͂��܂��B
		/// </summary>
		/// <param name="msg">�o�͂��镶����B</param>
		private void Print(string msg)
		{
			Console.WriteLine(msg);
		}
	}
}
