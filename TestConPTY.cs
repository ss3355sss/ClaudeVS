using System;
using System.Threading;

namespace ClaudeVS
{
    class TestConPTY
    {
        static void Main(string[] args)
        {
            var terminal = new ConPtyTerminal(30, 120);

            terminal.OutputReceived += (sender, output) =>
            {
            };

            terminal.ProcessExited += (sender, exitCode) =>
            {
            };

            bool success = terminal.Initialize();

            if (success)
            {
                Thread.Sleep(5000);

                terminal.WriteInput("dir\r\n");

                Thread.Sleep(2000);
            }

            terminal.Dispose();
        }
    }
}
