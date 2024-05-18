using System.Text;

namespace ExcelDiagramReplacer.Replacer
{
	internal class WordReplacer
	{
		private static string? _targetFile;

		private const string _extractPath = @"C:\replacer\xlsxFile";

		private const string _drawingPath = @"C:\replacer\xlsxFile\xl\drawings\";

		private const string _outputFile = @"C:\replacer\StringList.txt";

		private const string _prefixTag = @"<a:t>";

		private const string _suffixTag = @"</a:t>";

		private const string _nullWord = "[null]";

		private static Queue<string> _resultXmlLineQueue;


		public WordReplacer()
		{
			_resultXmlLineQueue = new Queue<string>();
		}

		/// <summary>
		/// 単語の置換処理
		/// </summary>
		public void doReplace()
		{
			if(!File.Exists(_outputFile))
			{
				MessageBox.Show("変換対象のテキストファイルがありません。\r\n C:/replacer/xlsxFile以下に配置してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			// drawing.xmlを探索
			for (int i = 1; i < int.MaxValue; i++)
			{
				StringBuilder sb = new StringBuilder();
				sb.Append(_drawingPath);
				sb.Append("drawing");
				sb.Append(i.ToString());
				sb.Append(".xml");

				// xml内の対象文字列の置換開始
				if (File.Exists(sb.ToString()))
				{
					ReplaceWords(sb.ToString());
				}
				else
				{
					return;
				}
			}
		}

		/// <summary>
		/// 各xmlの行に対して変換を行う
		/// </summary>
		/// <param name="xmlFile">drawing.xmlファイル</param>
		private void ReplaceWords(string xmlFile)
		{
			// XMLの読込み
			StreamReader sr = new StreamReader(xmlFile);
			string line = sr.ReadToEnd();
			sr.Close();

			// 各単語に対し、置換を行う
			foreach (var targetWord in System.IO.File.ReadLines(_outputFile))
			{
				// ::を区切りに、0要素目に日本語、1要素目に変換先文字列を入れる
				string[] wordsPair = targetWord.Split("::");

				// 対象文字列を含み、かつ翻訳済なら置換する
				if (line.Contains(wordsPair[0]) && !String.IsNullOrEmpty(wordsPair[1]))
				{
					// 対象外が明示的に指定されている場合、nullに変換する
					if (_nullWord.Equals(wordsPair[1]))
					{
						line = line.Replace(wordsPair[0], null);
					}
					else
					{
						// 置換処理
						line = line.Replace(wordsPair[0], wordsPair[1]);
					}

					// 書込み処理
					StreamWriter sw = new StreamWriter(xmlFile, false);
					sw.Write(line);
					sw.Close();
				}
			}
		}
	}
}
