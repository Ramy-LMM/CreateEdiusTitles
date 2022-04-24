using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Drawing;
using System.Windows.Forms;

namespace CreateEdiusTitles
{
	/// <summary>
	/// MainWindow.xaml の相互作用ロジック
	/// </summary>
	public partial class MainWindow : Window
	{
		//Ediusの設定値
		private int frameWidth = 1920;
		private int frameHeight = 1080;
		private string author = "hoge";
		private int fontSize = 60;

		public MainWindow()
		{
			InitializeComponent();
		}
		private void Create_Click(object sender, RoutedEventArgs e)
		{
			// ダイアログのインスタンスを生成
			var dialog = new Microsoft.Win32.OpenFileDialog();

			// ファイルの種類を設定
			dialog.Filter = "テキストファイル (*.txt)|*.txt";

			// ダイアログを表示する
			if (dialog.ShowDialog() == false)
			{
				return;
			}

			string dir = Path.GetDirectoryName(dialog.FileName) + "\\title\\";
			//titleフォルダの存在チェック
			if (!Directory.Exists(dir))
			{
				Directory.CreateDirectory(dir);
			}

			DateTime dt = DateTime.Now;
			try
			{   // streamReaderを開く
				using (StreamReader sr = new StreamReader(dialog.FileName, Encoding.GetEncoding("shift_jis")))
				{
					int count = 0;
					string text;
					while ((text = sr.ReadLine()) != null)
					{
						int length = text.Length;
						string xml = "";
						string save_path = dir + dt.ToString("yyyyMMdd") + string.Format("-{0:D5}", count) + ".etl";    //既存のetlファイルと区別するため、桁数を一つ増やす

						xml = "<?xml version=\"" + "1.0" + "\" encoding=\"" + "UTF-16" + "\"?>\n<!--Canopus Co., Ltd. Quick Titler-->\n<TitleDocument><Version>2.01</Version><Author>" + author + "</Author><CreateTime>" + dt.ToString("dd/MM/yyyy HH:mm:ss") + "</CreateTime><LastSavedETLPath>" + save_path + "</LastSavedETLPath><FrameSizeX>" + frameWidth + "</FrameSizeX><FrameSizeY>" + frameHeight + "</FrameSizeY><FrameAspectNumerator>" + frameWidth + "</FrameAspectNumerator><FrameAspectDenominator>" + frameWidth + "</FrameAspectDenominator><TitleType>0</TitleType><RollPages>1</RollPages><CrawlPages>1</CrawlPages><m_TitleObList><CCtsTextObject><CCtsTitleObject><m_sText>" + text + "</m_sText>";
						System.Drawing.Size nopadSize = new System.Drawing.Size();

						//描画先とするImageオブジェクトを作成する
						using (Bitmap canvas = new Bitmap(525, 350))
						{
							//ImageオブジェクトのGraphicsオブジェクトを作成する
							using (Graphics g = Graphics.FromImage(canvas))
							{
								//フォントオブジェクトの作成
								using (Font fnt = new Font("コーポレート・ロゴ（ラウンド） ver2 Bold", fontSize))
								{
									//NoPaddingにして、文字列を描画する
									TextRenderer.DrawText(g, text, fnt, new System.Drawing.Point(0, 50), Color.Black,
										TextFormatFlags.NoPadding);
									//大きさを計測
									nopadSize = TextRenderer.MeasureText(g, text, fnt,
										new System.Drawing.Size(1000, 1000), TextFormatFlags.NoPadding);
								}
							}
						}
						int width = nopadSize.Width + 34;    //11のエッジ(ハード幅6、ソフト幅5)＋なぜか12のオフセット(環境によって違うかも？)
						int x = (frameWidth - width) / 2 + 12;    //12のオフセット(環境によって違うかも？)
						int y = 478;    //<elements5>値＋3つ目の<TitlePointF>→<Y1>
						int element4 = x - 192; //<TitlePointF>→<X1>の数値

						//XMLファイルの中身の<m_sFontName>以下をコピー
						//適宜書き換えて下さい
						xml += "<m_sFontName>コーポレート・ロゴ（ラウンド） ver2 Bold</m_sFontName><m_sFillTextureFilePath></m_sFillTextureFilePath><m_sEdgeTextureFilePath></m_sEdgeTextureFilePath><m_sImageFilePath></m_sImageFilePath><dashStyle>19113544</dashStyle><startCap>19113708</startCap><endCap>197318</endCap><eSelect>0</eSelect><m_bRoundRectangle>0</m_bRoundRectangle><m_bTransformRatioMode>0</m_bTransformRatioMode><m_bRightAngled>0</m_bRightAngled><bMatrix>1</bMatrix><TitleFont><iSize>60</iSize><bVertical>0</bVertical><bBold>0</bBold><bItalic>0</bItalic><bUnderline>0</bUnderline><bAutoFontSize>0</bAutoFontSize><iCharSpace>0</iCharSpace><iLineSpace>0</iLineSpace><iAlignment>1</iAlignment></TitleFont><TitleFillColor><bEnableTexture>0</bEnableTexture><bEnable>0</bEnable><iColorNum>1</iColorNum><color0>-2469745</color0><color1>-16777216</color1><color2>-16777216</color2><color3>-16777216</color3><color4>-16777216</color4><color5>-16777216</color5><color6>-16777216</color6><fDirection>0.000000</fDirection></TitleFillColor><TitleEdge><bEnableTexture>0</bEnableTexture><bEnable>1</bEnable><iColorNum>1</iColorNum><color0>-1</color0><color1>-16777216</color1><color2>-16777216</color2><color3>-16777216</color3><color4>-16777216</color4><color5>-16777216</color5><color6>-16777216</color6><fWidth>6.000000</fWidth><fWidthSoft>5.000000</fWidthSoft><fDirection>0.000000</fDirection></TitleEdge><TitleShadow><bEnable>1</bEnable><fWidthHard>5.000000</fWidthHard><fWidthSoft>6.000000</fWidthSoft><iColorNum>1</iColorNum><color0>-15724528</color0><color1>-16777216</color1><color2>-16777216</color2><color3>-16777216</color3><color4>-16777216</color4><color5>-16777216</color5><color6>-16777216</color6><iOffsetX>5</iOffsetX><iOffsetY>6</iOffsetY><fDirection>0.000000</fDirection></TitleShadow><TitleBlur><bEnable>0</bEnable><iFillEdgeBlur>0</iFillEdgeBlur><iEdgeBlur>0</iEdgeBlur><iShadowBlur>0</iShadowBlur></TitleBlur><TitleEmboss><bEnable>0</bEnable><bOutside>0</bOutside><iEdgeHeight>3</iEdgeHeight><iFilter>2</iFilter><iLightX>1</iLightX><iLightY>2</iLightY><iLightZ>-3</iLightZ></TitleEmboss><elements0>1.000000</elements0><elements1>0.000000</elements1><elements2>0.000000</elements2><elements3>1.000000</elements3><elements4>" + element4 + ".000000</elements4><elements5>370.000000</elements5><TitleRectF><X>0.000000</X><Y>0.000000</Y><Width>0.000000</Width><Height>0.000000</Height></TitleRectF><TitleRectF><X>" + x + ".000000</X><Y>" + y + ".000000</Y><Width>" + nopadSize.Width + ".000000</Width><Height>" + nopadSize.Height + ".000000</Height></TitleRectF><TitlePointF><X1>192.000000</X1><Y1>108.000000</Y1><X2>0.000000</X2><Y2>0.000000</Y2></TitlePointF><TitleMultiPointF><m_nPointCount>2</m_nPointCount><X0>192.000000</X0><Y0>108.000000</Y0><X1>0.000000</X1><Y1>0.000000</Y1></TitleMultiPointF><TitleCharInfo><m_nCharCount>" + length + "</m_nCharCount><sFontName>" + length + ";コーポレート・ロゴ（ラウンド） ver2 Bold;</sFontName><bBold>" + length + ";0;</bBold><bItalic>" + length + ";0;</bItalic><bUnderline>" + length + ";0;</bUnderline><iCharSpace>" + length + ";0;</iCharSpace><iAveCharWidth>" + length + ";79;</iAveCharWidth><iLineSpace>" + length + ";0;</iLineSpace><iAveCharHeight>" + length + ";112;</iAveCharHeight><iSize>" + length + ";60;</iSize></TitleCharInfo></CCtsTitleObject></CCtsTextObject></m_TitleObList></TitleDocument>";

						using (StreamWriter sw = new StreamWriter(save_path, false, Encoding.GetEncoding("utf-16")))
						{
							sw.Write(xml);
						}
						count++;
					}

					this.FinishTextBlock.Text = "Finish!";
				}
			}
			catch (IOException ex)
			{
				Console.WriteLine("The file could not be read:");
				Console.WriteLine(ex.Message);
			}
		}
	}
}
