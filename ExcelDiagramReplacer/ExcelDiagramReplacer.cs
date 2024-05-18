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

		// 第一段階の加工が完了した状態の文字列リスト
		private static Queue<string> _linesOnFirstStep = new Queue<string>();

		// 最終的に出力すべき文字列
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
		/// 文字列一覧を出力するボタン押下時アクション
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button1_Click(object sender, EventArgs e)
		{
			try
			{
				// D&Dしたエクセルが存在しない時
				if (_targetFile == null)
				{
					MessageBox.Show("変換元エクセルが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
					// おわり
					return;
				}
				var destFile = _targetFile + ".target.zip";
				try
				{
					File.Copy(_targetFile, destFile);
				}
				catch (Exception ex2)
				{
					// すでにzipが存在する場合、何もしない
				}

				// zip解凍処理
				ZipFile.ExtractToDirectory(destFile, _extractPath);

				// 単語リストの取得
				DoWordCheck();

				// 単語リストをテキストに書込み
				WriteResult();

				// 完了メッセージ
				MessageBox.Show("文字列リストをC:/replacerに出力しました。", "出力完了！", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			}
			catch (Exception ex1)
			{
				MessageBox.Show("予期せぬエラーが発生しました。\r\n C/replacerを削除し、再度試してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine(ex1.Message);
			}

		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}

		/// <summary>
		/// ドラッグアンドドロップ時の処理
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void textBox1_DragDrop(object sender, DragEventArgs e)
		{
#pragma warning disable CS8602 // null 参照の可能性があるものの逆参照です。
			if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
#pragma warning restore CS8602 // null 参照の可能性があるものの逆参照です。

			string[] dragFilePathArr = (string[])e.Data.GetData(DataFormats.FileDrop, false);
			textBoxPath.Text = dragFilePathArr[0];
			_targetFile = dragFilePathArr[0];
		}

		/// <summary>
		/// ドラッグアンドドロップのドロップ時処理
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
		/// drawing.xmlを探索します
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

				// 検索開始
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
		/// 図形の文字列宣言の検索を行います
		/// </summary>
		/// <param name="file"></param>
		private void WordFinder(string file)
		{
			// 前方検索の開始
			foreach (string line in System.IO.File.ReadLines(file))
			{
				// line内に<a:t>が含まれているか確認
				SearchPrefix(line);
			}

			// 後方検索の開始
			SearchSuffix();
		}

		private void SearchPrefix(string line)
		{
			// prefixタグが含まれているか探索
			if (line.Contains(_prefixTag))
			{
				// line先頭からprefixタグとタグ含めた前方文字列を全て消去
				string lineOnFirstStep = line.Remove(0, line.IndexOf(_prefixTag) + 5);
				_linesOnFirstStep.Enqueue(lineOnFirstStep);
				// 削除済のlineに対し、再帰的にプレフィックス探索を行う
				SearchPrefix(lineOnFirstStep);
			}
		}

		private void SearchSuffix()
		{
			foreach (var line in _linesOnFirstStep)
			{
				if (line.Contains(_suffixTag))
				{
					// 初回登場のsuffixから、全体-suffixタグ登場以降の文字列を全て削除する
					string targetString = line.Remove(line.IndexOf(_suffixTag), line.Length - line.IndexOf(_suffixTag));

					// 同じ単語を重複してエンキューしないようにする
					if (!_targetStringQueue.Contains(targetString))
					{
						// 変換対象文字列を出力対象としてエンキュー
						_targetStringQueue.Enqueue(targetString);
					}
				}
			}
		}

		/// <summary>
		/// 変換対象をファイルに書き込みます。
		/// </summary>
		private void WriteResult()
		{
			// 出力文字列の作成
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
		/// 変換開始ボタン押下時イベント
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button2_Click(object sender, EventArgs e)
		{
			// 変換先文字列の置換作業
			Replacer.WordReplacer wr = new Replacer.WordReplacer();
			wr.doReplace();

			MessageBox.Show("文字の置き換えが完了しました。　\r\n C:/replacer直下にエクセルを出力しました。", "出力完了！", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

			ZipFile.CreateFromDirectory(_extractPath, _extractPath+"Output.xlsx");
		}
	}
}