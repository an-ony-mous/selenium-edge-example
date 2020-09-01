using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace SeleniumForEdge {
    /// <summary>
    /// WEBスクレイピングのサンプル
    /// </summary>
    public class Program {

        /// <summary>
        /// 実行
        /// </summary>
        /// <param name="args"></param>
        public static void Main(string[] args) {

            // 入力メッセージの表示
            Console.Write("商品検索文字を入力してください : ");

            // 入力文字取得
            var eValue = Console.ReadLine();

            var eWebScraping = new WebScraping();

            eWebScraping.Execute(eValue);
        }
    }

    /// <summary>
    /// Amazonで商品を検索するクラス
    /// </summary>
    public class WebScraping {
        /// <summary>
        /// 対象のURL
        /// </summary>
        private string TargetURL { get; } = "https://www.amazon.co.jp/";
        /// <summary>
        /// 対象のURL
        /// </summary>
        private string AppDir { get; }
        /// <summary>
        /// Webドライバ
        /// </summary>
        private IWebDriver WebDriver { set; get; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public WebScraping() {
            // 自分自身のディレクトリパス
            AppDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        /// <summary>
        /// Amazonから商品と価格を調べる
        /// </summary>
        /// <param name="pSearchValue"></param>
        public void Execute(string pSearchValue) {
            try {
                // ドライバが存在する場合は終了させる
                DriverKill();

                // デフォルトサービス設定(ドライバを指定するexeは"MicrosoftWebDriver.exe"固定みたい)
                var eService = EdgeDriverService.CreateDefaultService(AppDir, "MicrosoftWebDriver.exe");
                // コマンドプロンプト非表示
                eService.HideCommandPromptWindow = true;

                // オプション設定
                var eOptions = new EdgeOptions { PageLoadStrategy = (PageLoadStrategy)EdgePageLoadStrategy.Normal };

                // ドライバ生成
                WebDriver = new EdgeDriver(eService, eOptions);

                // URLセット
                WebDriver.Url = TargetURL;

                // 表示されるまで待つ
                WebDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(4);
                // ページタイトルに検索結果が含まれるまで
                ExecuteScript<IWebDriver>(WebDriver, $"document.getElementsByName('field-keywords')[0].value = \"{pSearchValue}\"");

                // 検索ボタンをクリック
                var eWebElement = WebDriver.FindElement(By.ClassName("nav-input"));
                eWebElement.Submit();

                // 検索結果のページが読み込み完了するまで待機する
                var eDriverWait = new WebDriverWait(WebDriver, TimeSpan.FromSeconds(20));
                eDriverWait.Until(ExpectedConditions.TitleContains(pSearchValue));

                // 商品ごとのXPath
                var eByPath = By.XPath("//span[contains(@class, 'a-size-base-plus a-color-base a-text-normal')]");
                // 商品一覧
                var eProducts = WebDriver.FindElements(eByPath);
                foreach (var eProduct in eProducts) {
                    // 商品から4つ上がって価格を調べる
                    IWebElement parent = eProduct.FindElement(By.XPath("../../../.."));
                    // 価格の要素を取得する
                    var ePrice = parent.FindElement(By.ClassName("a-price-whole"));
                    // コンソール出力
                    Console.WriteLine($"商品名:{eProduct.Text} 価格:{ePrice.Text}");
                }
            }
            catch (Exception eException) {
                // コンソール出力
                Console.WriteLine(eException.Message);
                // デバッグログ出力
                Debug.WriteLine(eException.Message);
            }
            finally {
                // ドライバの終了
                WebDriver.Quit();
            }
        }

        /// <summary>
        /// 既に存在しているドライバを終了させる
        /// </summary>
        public void DriverKill() {
            // 既に存在しているドライバを終了させる
            foreach (var eProcess in Process.GetProcessesByName("MicrosoftWebDriver")) {
                // 作成したプロセスが終了していないか確認する
                if (!eProcess.HasExited) {
                    // 作成したプロセスを強制終了する
                    eProcess.Kill();
                }
                eProcess.Close();
            }
        }

        /// <summary>
        /// JavaScriptを実行させる
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="pWebDriver"></param>
        /// <param name="pScript"></param>
        /// <returns></returns>
        public T ExecuteScript<T>(IWebDriver pWebDriver, string pScript) {
            if (pWebDriver is IJavaScriptExecutor) {
                return (T)((IJavaScriptExecutor)pWebDriver).ExecuteScript(pScript);
            }
            else {
                throw new WebDriverException();
            }
        }
    }
}
