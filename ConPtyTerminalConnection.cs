namespace ClaudeVS
{
    using System;
    using System.Diagnostics;
    using System.Text;
    using System.Threading;
    using Microsoft.Terminal.Wpf;

    /// <summary>
    /// Custom TerminalConnection that bridges ConPtyTerminal with Microsoft.Terminal.Wpf.TerminalControl
    /// </summary>
    public class ConPtyTerminalConnection : ITerminalConnection
    {
        private readonly ConPtyTerminal conPtyTerminal;
        private readonly ManualResetEventSlim connectionReadyEvent = new ManualResetEventSlim(false);

        public ConPtyTerminalConnection(ConPtyTerminal terminal)
        {
            if (terminal == null)
            {
                Debug.WriteLine("ArgumentNullException in ConPtyTerminalConnection constructor: terminal is null");
                throw new ArgumentNullException(nameof(terminal));
            }
            conPtyTerminal = terminal;

            conPtyTerminal.OutputReceived += (sender, output) =>
            {
                if (terminalOutputEvent != null)
                {
                    terminalOutputEvent.Invoke(this, new TerminalOutputEventArgs(output));
                }
            };

            conPtyTerminal.ProcessExited += (sender, exitCode) =>
            {
                Closed?.Invoke(this, EventArgs.Empty);
            };
        }

        public event EventHandler<TerminalOutputEventArgs> TerminalOutput
        {
            add
            {
                terminalOutputEvent += value;
                connectionReadyEvent.Set();
            }
            remove
            {
                terminalOutputEvent -= value;
            }
        }

        private event EventHandler<TerminalOutputEventArgs> terminalOutputEvent;

        public event EventHandler Closed;

        public void WaitForConnectionReady()
        {
            bool ready = connectionReadyEvent.Wait(TimeSpan.FromSeconds(5));
        }

        public void WriteInput(string data)
        {
            try
            {
                if (conPtyTerminal != null && conPtyTerminal.IsRunning)
                    conPtyTerminal.WriteInput(data);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in WriteInput: {ex}");
            }
        }

        public void Resize(uint rows, uint columns)
        {
            try
            {
                if (conPtyTerminal != null && conPtyTerminal.IsRunning)
                {
                    conPtyTerminal.Resize((ushort)rows, (ushort)columns);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in Resize: {ex}");
            }
        }

        public void Close()
        {
            try
            {
                conPtyTerminal?.Dispose();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in Close: {ex}");
            }
        }

        public void Start()
        {
        }
    }
}
