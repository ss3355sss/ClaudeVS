namespace ClaudeVS
{
	using System;
	using System.ComponentModel.Design;
	using System.Diagnostics;
	using System.Runtime.InteropServices;
	using System.Text;
	using EnvDTE;
	using EnvDTE80;
	using EnvDTE90a;
	using Microsoft.VisualStudio;
	using Microsoft.VisualStudio.Debugger.Interop;
	using Microsoft.VisualStudio.Shell;
	using Microsoft.VisualStudio.Shell.Interop;
	using Task = System.Threading.Tasks.Task;

	internal sealed class SendDebuggerExceptionCommand : IDebugEventCallback2
	{
		public const int CommandId = 0x0104;

		public static readonly Guid CommandSet = new Guid("a7c8e9d0-1234-5678-9abc-def012345678");

		private readonly AsyncPackage package;

		private static string _lastExceptionType;
		private static string _lastExceptionName;
		private static uint _lastExceptionCode;
		private static string _lastExceptionDescription;
		private static DebuggerEvents _debuggerEvents;

		private SendDebuggerExceptionCommand(AsyncPackage package, OleMenuCommandService commandService)
		{
			if (package == null)
			{
				Debug.WriteLine("ArgumentNullException in SendDebuggerExceptionCommand constructor: package is null");
				throw new ArgumentNullException(nameof(package));
			}
			if (commandService == null)
			{
				Debug.WriteLine("ArgumentNullException in SendDebuggerExceptionCommand constructor: commandService is null");
				throw new ArgumentNullException(nameof(commandService));
			}
			this.package = package;

			var menuCommandID = new CommandID(CommandSet, CommandId);
			var menuItem = new OleMenuCommand(this.Execute, menuCommandID);
			menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
			commandService.AddCommand(menuItem);

			SubscribeToDebuggerEvents();
		}

		private void SubscribeToDebuggerEvents()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				IVsDebugger debugger = Package.GetGlobalService(typeof(SVsShellDebugger)) as IVsDebugger;
				if (debugger != null)
				{
					debugger.AdviseDebugEventCallback(this);
				}

				DTE2 dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
				if (dte != null)
				{
					_debuggerEvents = dte.Events.DebuggerEvents;
					_debuggerEvents.OnExceptionThrown += OnExceptionThrown;
					_debuggerEvents.OnEnterBreakMode += OnEnterBreakMode;
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Failed to subscribe to debugger events: {ex.Message}");
			}
		}

		public int Event(IDebugEngine2 pEngine, IDebugProcess2 pProcess, IDebugProgram2 pProgram, 
			IDebugThread2 pThread, IDebugEvent2 pEvent, ref Guid riidEvent, uint dwAttrib)
		{
			try
			{
				if (pEvent is IDebugExceptionEvent2 exceptionEvent)
				{
					EXCEPTION_INFO[] exInfo = new EXCEPTION_INFO[1];
					if (exceptionEvent.GetException(exInfo) == VSConstants.S_OK)
					{
						_lastExceptionName = exInfo[0].bstrExceptionName;
						_lastExceptionCode = exInfo[0].dwCode;
					}

					if (exceptionEvent.GetExceptionDescription(out string description) == VSConstants.S_OK)
					{
						_lastExceptionDescription = description;
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Error in IDebugEventCallback2.Event: {ex.Message}");
			}

			return VSConstants.S_OK;
		}

		private void OnExceptionThrown(string exceptionType, string name, int code, string description, ref dbgExceptionAction exceptionAction)
		{
			_lastExceptionType = exceptionType;
			_lastExceptionName = name;
			_lastExceptionCode = (uint)code;
			_lastExceptionDescription = description;
		}

		private void OnEnterBreakMode(dbgEventReason reason, ref dbgExecutionAction executionAction)
		{
			if (reason != dbgEventReason.dbgEventReasonExceptionThrown && reason != dbgEventReason.dbgEventReasonExceptionNotHandled)
			{
				_lastExceptionType = null;
				_lastExceptionName = null;
				_lastExceptionCode = 0;
				_lastExceptionDescription = null;
			}
		}

		public static SendDebuggerExceptionCommand Instance
		{
			get;
			private set;
		}

		private IAsyncServiceProvider ServiceProvider
		{
			get { return this.package; }
		}

		public static async Task InitializeAsync(AsyncPackage package)
		{
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new SendDebuggerExceptionCommand(package, commandService);
		}

		private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (sender is OleMenuCommand menuCommand)
			{
				try
				{
					DTE2 dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
					if (dte?.Debugger != null)
					{
						menuCommand.Visible = dte.Debugger.CurrentMode == dbgDebugMode.dbgBreakMode;
					}
					else
					{
						menuCommand.Visible = false;
					}
				}
				catch
				{
					menuCommand.Visible = false;
				}
			}
		}

		private void Execute(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			try
			{
				DTE2 dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
				if (dte?.Debugger == null || dte.Debugger.CurrentMode != dbgDebugMode.dbgBreakMode)
				{
					return;
				}

				StringBuilder message = new StringBuilder();
				message.AppendLine("DEBUGGER EXCEPTION/ERROR INFO:");
				message.AppendLine();

				string exceptionMessage = GetExceptionMessage(dte);

				if (!string.IsNullOrEmpty(exceptionMessage))
				{
					message.AppendLine("=== EXCEPTION ===");
					message.AppendLine(exceptionMessage);
					message.AppendLine();
				}

				Expression lastException = null;
				try
				{
					lastException = dte.Debugger.GetExpression("$exception", false, 1000);
				}
				catch { }

				if (lastException != null && lastException.IsValidValue)
				{
					if (lastException.Type != null && lastException.Type.Contains("System."))
					{
						message.AppendLine("=== MANAGED EXCEPTION DETAILS ===");
						message.AppendLine($"Type: {lastException.Type}");

						try
						{
							var stackTrace = dte.Debugger.GetExpression("$exception.StackTrace", false, 1000);
							if (stackTrace != null && stackTrace.IsValidValue)
							{
								message.AppendLine();
								message.AppendLine("Stack Trace:");
								message.AppendLine(stackTrace.Value);
							}
						}
						catch { }

						try
						{
							var innerEx = dte.Debugger.GetExpression("$exception.InnerException", false, 1000);
							if (innerEx != null && innerEx.IsValidValue && !innerEx.Value.Contains("null"))
							{
								message.AppendLine();
								message.AppendLine($"Inner Exception: {innerEx.Value}");
							}
						}
						catch { }
						message.AppendLine();
					}
				}

				if (dte.Debugger.LastBreakReason == dbgEventReason.dbgEventReasonExceptionThrown ||
					dte.Debugger.LastBreakReason == dbgEventReason.dbgEventReasonExceptionNotHandled)
				{
					message.AppendLine("=== BREAK REASON ===");
					message.AppendLine($"Reason: {dte.Debugger.LastBreakReason}");
					message.AppendLine();
				}

				if (dte.Debugger.CurrentThread != null && dte.Debugger.CurrentThread.StackFrames != null)
				{
					try
					{
						message.AppendLine("=== CALL STACK ===");
						int frameCount = 0;
						foreach (EnvDTE.StackFrame frame in dte.Debugger.CurrentThread.StackFrames)
						{
							if (frameCount >= 20)
								break;
							
							string frameInfo = $"{frameCount}: {frame.FunctionName}";
							
							try
							{
								string fileInfo = GetStackFrameFileInfo(frame);
								if (!string.IsNullOrEmpty(fileInfo))
									frameInfo += $" - {fileInfo}";
							}
							catch { }
							
							message.AppendLine(frameInfo);
							frameCount++;
						}
						message.AppendLine();
					}
					catch { }
				}

				if (dte.Debugger.CurrentStackFrame != null)
				{
					message.AppendLine();
					message.AppendLine("=== CURRENT STACK FRAME ===");
					message.AppendLine($"Function: {dte.Debugger.CurrentStackFrame.FunctionName}");
					message.AppendLine($"Language: {dte.Debugger.CurrentStackFrame.Language}");
					
					try
					{
						var localsExpr = dte.Debugger.CurrentStackFrame.Locals;
						if (localsExpr != null)
						{
							foreach (EnvDTE.Expression local in localsExpr)
							{
								if (local.Name == "$T0")
								{
									string sourceInfo = local.Value;
									if (!string.IsNullOrEmpty(sourceInfo) && sourceInfo.Contains(":"))
									{
										message.AppendLine($"Source: {sourceInfo}");
									}
									break;
								}
							}
						}
					}
					catch { }
					
					if (dte.Debugger.CurrentStackFrame.Module != null)
					{
						message.AppendLine($"Module: {dte.Debugger.CurrentStackFrame.Module}");
					}
				}

				if (dte.Debugger.BreakpointLastHit != null)
				{
					message.AppendLine();
					message.AppendLine("=== BREAKPOINT INFO ===");
					message.AppendLine($"File: {dte.Debugger.BreakpointLastHit.File}");
					message.AppendLine($"Line: {dte.Debugger.BreakpointLastHit.FileLine}");
					message.AppendLine($"Condition: {dte.Debugger.BreakpointLastHit.Condition}");
				}

				if (dte.Debugger.CurrentProcess != null)
				{
					message.AppendLine();
					message.AppendLine("=== PROCESS INFO ===");
					message.AppendLine($"Name: {dte.Debugger.CurrentProcess.Name}");
					message.AppendLine($"Process ID: {dte.Debugger.CurrentProcess.ProcessID}");

					if (dte.Debugger.CurrentThread != null)
					{
						message.AppendLine($"Thread ID: {dte.Debugger.CurrentThread.ID}");
						message.AppendLine($"Thread Name: {dte.Debugger.CurrentThread.Name}");
					}
				}

				string outputWindowText = GetLastOutputWindowLines(dte, 10);
				if (!string.IsNullOrEmpty(outputWindowText))
				{
					message.AppendLine();
					message.AppendLine("=== OUTPUT WINDOW (Last 10 lines) ===");
					message.AppendLine(outputWindowText);
				}

				ToolWindowPane window = this.package.FindToolWindow(typeof(ClaudeTerminal), 0, true);
				if (window?.Frame is IVsWindowFrame windowFrame)
				{
					Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());

					if (window.Content is ClaudeTerminalControl control)
					{
						control.SendToClaude("\x1b[200~" + message.ToString() + "\x1b[201~", true);
						control.FocusTerminal();
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Exception in SendDebuggerExceptionCommand Execute: {ex}");
			}
		}

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		private static extern int FormatMessage(
			int dwFlags,
			IntPtr lpSource,
			uint dwMessageId,
			int dwLanguageId,
			StringBuilder lpBuffer,
			int nSize,
			IntPtr Arguments);

		private const int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

		private string GetExceptionMessage(DTE2 dte)
		{
			if (!string.IsNullOrEmpty(_lastExceptionName) || _lastExceptionCode != 0)
			{
				StringBuilder sb = new StringBuilder();
				
				if (!string.IsNullOrEmpty(_lastExceptionType))
					sb.AppendLine($"Type: {_lastExceptionType}");
				
				if (!string.IsNullOrEmpty(_lastExceptionName))
					sb.AppendLine($"Name: {_lastExceptionName}");
				
				if (_lastExceptionCode != 0)
				{
					sb.AppendLine($"Code: 0x{_lastExceptionCode:X8}");
					string decoded = DecodeSEH(_lastExceptionCode);
					sb.AppendLine($"Description: {decoded}");
				}
				
				if (!string.IsNullOrEmpty(_lastExceptionDescription))
				{
					sb.AppendLine($"Details: {_lastExceptionDescription}");
				}
				
				return sb.ToString().TrimEnd();
			}
			
			try
			{
				var cppMsg = dte.Debugger.GetExpression("((std::exception*)$exception)->what()", false, 1000);

				if (cppMsg != null && cppMsg.IsValidValue && !cppMsg.Value.ToLower().Contains("error") && !cppMsg.Value.Contains("undefined"))
				{
					string msg = cppMsg.Value.Trim('"', ' ');
					if (!string.IsNullOrEmpty(msg))
						return $"C++ Exception: {msg}";
				}
			}
			catch { }

			try
			{
				var managedEx = dte.Debugger.GetExpression("$exception.Message", false, 1000);
				if (managedEx != null && managedEx.IsValidValue)
				{
					return $"Managed Exception: {managedEx.Value.Trim('"')}";
				}
			}
			catch { }

			return null;
		}

		private string DecodeSEH(uint errorCode)
		{
			StringBuilder buffer = new StringBuilder(1024);
			int result = FormatMessage(
				FORMAT_MESSAGE_FROM_SYSTEM,
				IntPtr.Zero,
				errorCode,
				0,
				buffer,
				buffer.Capacity,
				IntPtr.Zero);

			if (result > 0)
			{
				return buffer.ToString().Trim();
			}

			switch (errorCode)
			{
				case 0xC0000005: return "Access Violation (Memory Read/Write Error)";
				case 0xC00000FD: return "Stack Overflow";
				case 0xC0000094: return "Integer Division by Zero";
				case 0xC000001D: return "Illegal Instruction";
				case 0xC000008C: return "Array Bounds Exceeded";
				case 0xC0000409: return "Stack Buffer Overrun";
				case 0x80000003: return "Breakpoint";
				case 0xC0000008: return "Invalid Handle";
				case 0xC000013A: return "Control-C Exit";
				default: return $"System Exception: 0x{errorCode:X8}";
			}
		}

		private string GetStackFrameFileInfo(EnvDTE.StackFrame frame)
		{
			try
			{
				var frame2 = frame as StackFrame2;
				if (frame2 != null)
				{
					string fileName = frame2.FileName;
					uint lineNumber = frame2.LineNumber;
					
					if (!string.IsNullOrEmpty(fileName) && lineNumber > 0)
					{
						string shortName = System.IO.Path.GetFileName(fileName);
						return $"{shortName}:{lineNumber}";
					}
				}
			}
			catch { }

			return null;
		}

		private string GetLastOutputWindowLines(DTE2 dte, int lineCount)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			try
			{
				OutputWindow outputWindow = dte.ToolWindows.OutputWindow;
				if (outputWindow == null || outputWindow.OutputWindowPanes == null)
					return null;

				OutputWindowPane debugPane = null;
				foreach (OutputWindowPane pane in outputWindow.OutputWindowPanes)
				{
					if (pane.Name == "Debug")
					{
						debugPane = pane;
						break;
					}
				}

				if (debugPane == null)
					return null;

				TextDocument textDoc = debugPane.TextDocument;
				if (textDoc == null)
					return null;

				EditPoint startPoint = textDoc.StartPoint.CreateEditPoint();
				EditPoint endPoint = textDoc.EndPoint.CreateEditPoint();
				string allText = startPoint.GetText(endPoint);

				if (string.IsNullOrEmpty(allText))
					return null;

				string[] allLines = allText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
				var nonEmptyLines = new System.Collections.Generic.List<string>();
				for (int i = allLines.Length - 1; i >= 0 && nonEmptyLines.Count < lineCount; i--)
				{
					if (!string.IsNullOrWhiteSpace(allLines[i]))
						nonEmptyLines.Insert(0, allLines[i]);
				}

				if (nonEmptyLines.Count == 0)
					return null;

				return string.Join(Environment.NewLine, nonEmptyLines);
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Failed to get output window text: {ex.Message}");
				return null;
			}
		}
	}
}
