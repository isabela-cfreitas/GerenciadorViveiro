using System.IO;
using System.Reflection;

namespace GerenciadorViveiro;

public static class Log{
    static Log(){
        File.Delete(path + "\\" + "log.txt");
    }
    private static string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!;
    public static void WriteLine(string Text){
        //path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        using StreamWriter S = File.AppendText(path + "\\" + "log.txt");
        S.Write(Text + '\n');
    }
}