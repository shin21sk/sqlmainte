using SqlMainte.Forms;

namespace SqlMainte;

static class Program
{
    [STAThread]
    static void Main()
    {
        // 未処理例外をメッセージボックスで表示（無言クラッシュ防止）
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) =>
            MessageBox.Show($"予期しないエラーが発生しました。\n\n{e.Exception.Message}\n\n{e.Exception.StackTrace}",
                "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            MessageBox.Show($"致命的なエラーが発生しました。\n\n{e.ExceptionObject}",
                "致命的エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}