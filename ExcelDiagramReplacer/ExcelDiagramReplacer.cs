using System.IO.Compression;
using System.Text;

namespace ExcelDiagramReplacer
{
	public partial class ExcelDiagramReplacer : Form
	{

		private static string? _targetFile;

		private const string _extractPath = @"C:\replacer\xlsxFile";

		private const string _drawingPath = @"C:\replacer\xlsxFile\xl\drawings\";

		private const string _outputFile = @"C:\replacer\StringList.txt";

		private const string _prefixTag = @"<a:t>";

		private const string _suffixTag = @"</a:t>";

		// ���i�K�̉��H������������Ԃ̕����񃊃X�g
		private static Queue<string> _linesOnFirstStep = new Queue<string>();

		// �ŏI�I�ɏo�͂��ׂ�������
		private static Queue<string> _targetStringQueue = new Queue<string>();

		/// <summary>
		/// constractor
		/// </summary>
		public ExcelDiagramReplacer()
		{
			InitializeComponent();
			_targetStringQueue = new Queue<string>();
			_linesOnFirstStep = new Queue<string>();
		}

		/// <summary>
		/// ������ꗗ���o�͂���{�^���������A�N�V����
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button1_Click(object sender, EventArgs e)
		{
			try
			{
				// D&D�����G�N�Z�������݂��Ȃ���
				if (_targetFile == null)
				{
					MessageBox.Show("�ϊ����G�N�Z�������݂��܂���B", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
					// �����
					return;
				}
				var destFile = _targetFile + ".target.zip";
				try
				{
					File.Copy(_targetFile, destFile);
				}
				catch (Exception ex2)
				{
					// ���ł�zip�����݂���ꍇ�A�������Ȃ�
				}

				// zip�𓀏���
				ZipFile.ExtractToDirectory(destFile, _extractPath);

				// �P�ꃊ�X�g�̎擾
				DoWordCheck();

				// �P�ꃊ�X�g���e�L�X�g�ɏ�����
				WriteResult();

				// �������b�Z�[�W
				MessageBox.Show("�����񃊃X�g��C:/replacer�ɏo�͂��܂����B", "�o�͊����I", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			}
			catch (Exception ex1)
			{
				MessageBox.Show("�\�����ʃG���[���������܂����B\r\n C/replacer���폜���A�ēx�����Ă��������B", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine(ex1.Message);
			}

		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}

		/// <summary>
		/// �h���b�O�A���h�h���b�v���̏���
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void textBox1_DragDrop(object sender, DragEventArgs e)
		{
#pragma warning disable CS8602 // null �Q�Ƃ̉\����������̂̋t�Q�Ƃł��B
			if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
#pragma warning restore CS8602 // null �Q�Ƃ̉\����������̂̋t�Q�Ƃł��B

			string[] dragFilePathArr = (string[])e.Data.GetData(DataFormats.FileDrop, false);
			textBoxPath.Text = dragFilePathArr[0];
			_targetFile = dragFilePathArr[0];
		}

		/// <summary>
		/// �h���b�O�A���h�h���b�v�̃h���b�v������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void textBox1_DragEnter(object sender, DragEventArgs e)
		{
			e.Effect = DragDropEffects.All;
		}

		private void ChangeDirectoryAttribute(string directory)
		{
			DirectoryInfo di = new DirectoryInfo(directory);
			if ((di.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
			{
				di.Attributes = FileAttributes.Normal;
			}
		}

		/// <summary>
		/// drawing.xml��T�����܂�
		/// </summary>
		private void DoWordCheck()
		{
			for (int i = 1; i < int.MaxValue; i++)
			{
				StringBuilder sb = new StringBuilder();
				sb.Append(_drawingPath);
				sb.Append("drawing");
				sb.Append(i.ToString());
				sb.Append(".xml");

				// �����J�n
				if (File.Exists(sb.ToString()))
				{
					WordFinder(sb.ToString());
				}
				else
				{
					return;
				}
			}
		}

		/// <summary>
		/// �}�`�̕�����錾�̌������s���܂�
		/// </summary>
		/// <param name="file"></param>
		private void WordFinder(string file)
		{
			// �O�������̊J�n
			foreach (string line in System.IO.File.ReadLines(file))
			{
				// line����<a:t>���܂܂�Ă��邩�m�F
				SearchPrefix(line);
			}

			// ��������̊J�n
			SearchSuffix();
		}

		private void SearchPrefix(string line)
		{
			// prefix�^�O���܂܂�Ă��邩�T��
			if (line.Contains(_prefixTag))
			{
				// line�擪����prefix�^�O�ƃ^�O�܂߂��O���������S�ď���
				string lineOnFirstStep = line.Remove(0, line.IndexOf(_prefixTag) + 5);
				_linesOnFirstStep.Enqueue(lineOnFirstStep);
				// �폜�ς�line�ɑ΂��A�ċA�I�Ƀv���t�B�b�N�X�T�����s��
				SearchPrefix(lineOnFirstStep);
			}
		}

		private void SearchSuffix()
		{
			foreach (var line in _linesOnFirstStep)
			{
				if (line.Contains(_suffixTag))
				{
					// ����o���suffix����A�S��-suffix�^�O�o��ȍ~�̕������S�č폜����
					string targetString = line.Remove(line.IndexOf(_suffixTag), line.Length - line.IndexOf(_suffixTag));

					// �����P����d�����ăG���L���[���Ȃ��悤�ɂ���
					if (!_targetStringQueue.Contains(targetString))
					{
						// �ϊ��Ώە�������o�͑ΏۂƂ��ăG���L���[
						_targetStringQueue.Enqueue(targetString);
					}
				}
			}
		}

		/// <summary>
		/// �ϊ��Ώۂ��t�@�C���ɏ������݂܂��B
		/// </summary>
		private void WriteResult()
		{
			// �o�͕�����̍쐬
			StringBuilder sb = new StringBuilder();
			foreach (var str in _targetStringQueue)
			{
				sb.Append(str);
				sb.Append("::");
				sb.Append("\r\n");
			}
			File.WriteAllText(_outputFile, sb.ToString());
		}

		/// <summary>
		/// �ϊ��J�n�{�^���������C�x���g
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button2_Click(object sender, EventArgs e)
		{
			// �ϊ��敶����̒u�����
			Replacer.WordReplacer wr = new Replacer.WordReplacer();
			wr.doReplace();

			MessageBox.Show("�����̒u���������������܂����B�@\r\n C:/replacer�����ɃG�N�Z�����o�͂��܂����B", "�o�͊����I", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

			ZipFile.CreateFromDirectory(_extractPath, _extractPath+"Output.xlsx");
		}
	}
}