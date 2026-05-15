namespace ClaudeVS
{
	using System;
	using System.Collections.Generic;
	using System.Diagnostics;
	using System.IO;
	using System.Net;
	using System.Text;
	using System.Web.Script.Serialization;
	using System.Windows;
	using EnvDTE;
	using EnvDTE80;
	using Microsoft.VisualStudio.Shell;

	internal class McpServer : IDisposable
	{
		private HttpListener listener;
		private System.Threading.Thread listenerThread;
		private volatile bool running;
		private readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
		private int port;

		public int Port => port;

		public bool Start()
		{
			for (int p = 4322; p < 4332; p++)
			{
				try
				{
					var l = new HttpListener();
					l.Prefixes.Add($"http://127.0.0.1:{p}/");
					l.Start();
					listener = l;
					port = p;
					running = true;
					listenerThread = new System.Threading.Thread(ListenLoop) { IsBackground = true, Name = "ClaudeVS MCP" };
					listenerThread.Start();
					Debug.WriteLine($"ClaudeVS MCP server started on port {p}");
					return true;
				}
				catch
				{
					continue;
				}
			}
			Debug.WriteLine("ClaudeVS MCP server failed to start on any port");
			return false;
		}

		private void ListenLoop()
		{
			while (running)
			{
				try
				{
					var context = listener.GetContext();
					System.Threading.ThreadPool.QueueUserWorkItem(_ => HandleRequest(context));
				}
				catch (HttpListenerException) when (!running) { }
				catch (Exception ex)
				{
					Debug.WriteLine($"MCP listener error: {ex.Message}");
				}
			}
		}

		private void HandleRequest(HttpListenerContext context)
		{
			try
			{
				context.Response.Headers.Add("Access-Control-Allow-Origin", "*");
				context.Response.Headers.Add("Access-Control-Allow-Methods", "POST, GET, OPTIONS");
				context.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Mcp-Session-Id");

				if (context.Request.HttpMethod == "OPTIONS")
				{
					context.Response.StatusCode = 204;
					context.Response.Close();
					return;
				}

				if (context.Request.HttpMethod == "GET")
				{
					SendJsonResponse(context, new Dictionary<string, object>
					{
						["name"] = "ClaudeVS",
						["version"] = "1.0",
						["status"] = "running"
					});
					return;
				}

				if (context.Request.HttpMethod != "POST")
				{
					context.Response.StatusCode = 405;
					context.Response.Close();
					return;
				}

				string body;
				using (var reader = new StreamReader(context.Request.InputStream, Encoding.UTF8))
					body = reader.ReadToEnd();

				var request = json.Deserialize<Dictionary<string, object>>(body);
				string method = request.ContainsKey("method") ? request["method"] as string : null;
				object id = request.ContainsKey("id") ? request["id"] : null;

				if (method == null)
				{
					context.Response.StatusCode = 400;
					context.Response.Close();
					return;
				}

				if (id == null)
				{
					context.Response.StatusCode = 202;
					context.Response.Close();
					return;
				}

				object result = HandleMethod(method, request);

				bool isError = result is Dictionary<string, object> d && d.ContainsKey("code") && d.ContainsKey("message");
				var rpcResponse = new Dictionary<string, object>
				{
					["jsonrpc"] = "2.0",
					["id"] = id
				};
				if (isError)
					rpcResponse["error"] = result;
				else
					rpcResponse["result"] = result;

				SendJsonResponse(context, rpcResponse);
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"MCP request error: {ex.Message}");
				try { context.Response.StatusCode = 500; context.Response.Close(); }
				catch { }
			}
		}

		private object HandleMethod(string method, Dictionary<string, object> request)
		{
			switch (method)
			{
				case "initialize":
					return new Dictionary<string, object>
					{
						["protocolVersion"] = "2025-11-25",
						["capabilities"] = new Dictionary<string, object>
						{
							["tools"] = new Dictionary<string, object>()
						},
						["serverInfo"] = new Dictionary<string, object>
						{
							["name"] = "ClaudeVS",
							["version"] = "1.0"
						}
					};

				case "tools/list":
					return new Dictionary<string, object>
					{
						["tools"] = GetToolDefinitions()
					};

				case "tools/call":
					return HandleToolCall(request);

				default:
					return new Dictionary<string, object>
					{
						["code"] = -32601,
						["message"] = $"Unknown method: {method}"
					};
			}
		}

		private object[] GetToolDefinitions()
		{
			return new object[]
			{
				new Dictionary<string, object>
				{
					["name"] = "vs_getEditorContext",
					["description"] = "Get the current Visual Studio editor state: active file path (relative to solution), cursor line/column, selected text, and surrounding source code (±15 lines around cursor).",
					["inputSchema"] = new Dictionary<string, object>
					{
						["type"] = "object",
						["properties"] = new Dictionary<string, object>()
					}
				},
				new Dictionary<string, object>
				{
					["name"] = "vs_getDiagnostics",
					["description"] = "Get all errors and warnings from the Visual Studio Error List for the current solution. Returns file, line, description for each item.",
					["inputSchema"] = new Dictionary<string, object>
					{
						["type"] = "object",
						["properties"] = new Dictionary<string, object>()
					}
				},
				new Dictionary<string, object>
				{
					["name"] = "vs_getOpenFiles",
					["description"] = "Get list of all currently open documents in Visual Studio with their file paths.",
					["inputSchema"] = new Dictionary<string, object>
					{
						["type"] = "object",
						["properties"] = new Dictionary<string, object>()
					}
				}
			};
		}

		private object HandleToolCall(Dictionary<string, object> request)
		{
			var paramsObj = request.ContainsKey("params") ? request["params"] as Dictionary<string, object> : null;
			string toolName = paramsObj?.ContainsKey("name") == true ? paramsObj["name"] as string : null;

			if (toolName == null)
				return new Dictionary<string, object> { ["code"] = -32602, ["message"] = "Missing tool name" };

			string resultText;
			try
			{
				resultText = (string)Application.Current.Dispatcher.Invoke(new Func<string>(() =>
				{
					ThreadHelper.ThrowIfNotOnUIThread();
					DTE2 dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
					if (dte == null)
						return "Visual Studio DTE not available";

					switch (toolName)
					{
						case "vs_getEditorContext": return GetEditorContext(dte);
						case "vs_getDiagnostics": return GetDiagnostics(dte);
						case "vs_getOpenFiles": return GetOpenFiles(dte);
						default: return $"Unknown tool: {toolName}";
					}
				}));
			}
			catch (Exception ex)
			{
				resultText = $"Error: {ex.Message}";
			}

			return new Dictionary<string, object>
			{
				["content"] = new object[]
				{
					new Dictionary<string, object>
					{
						["type"] = "text",
						["text"] = resultText
					}
				}
			};
		}

		private string GetEditorContext(DTE2 dte)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (dte.ActiveDocument == null)
				return "No active document";

			string filePath = dte.ActiveDocument.FullName;
			TextSelection selection = dte.ActiveDocument.Selection as TextSelection;
			int line = selection?.CurrentLine ?? 1;
			int col = selection?.CurrentColumn ?? 1;
			string selectedText = selection?.Text;

			string solutionDir = null;
			if (dte.Solution != null && !string.IsNullOrEmpty(dte.Solution.FullName))
				solutionDir = Path.GetDirectoryName(dte.Solution.FullName);

			string relativePath = filePath;
			if (!string.IsNullOrEmpty(solutionDir) && filePath.StartsWith(solutionDir, StringComparison.OrdinalIgnoreCase))
				relativePath = filePath.Substring(solutionDir.Length).TrimStart('\\', '/');

			string ext = Path.GetExtension(filePath).TrimStart('.');

			var sb = new StringBuilder();
			sb.AppendLine($"File: {relativePath}");
			sb.AppendLine($"Cursor: line {line}, column {col}");

			if (!string.IsNullOrEmpty(selectedText))
			{
				sb.AppendLine();
				sb.AppendLine("Selected text:");
				sb.AppendLine($"```{ext}");
				sb.AppendLine(selectedText);
				sb.AppendLine("```");
			}

			try
			{
				TextDocument textDoc = dte.ActiveDocument.Object("TextDocument") as TextDocument;
				if (textDoc != null)
				{
					int startLine = Math.Max(1, line - 15);
					int endLine = Math.Min(textDoc.EndPoint.Line, line + 15);

					EditPoint startPoint = textDoc.CreateEditPoint();
					startPoint.MoveToLineAndOffset(startLine, 1);
					EditPoint endPoint = textDoc.CreateEditPoint();
					endPoint.MoveToLineAndOffset(endLine, 1);
					endPoint.EndOfLine();

					string code = startPoint.GetText(endPoint);

					sb.AppendLine();
					sb.AppendLine($"Surrounding code (lines {startLine}-{endLine}, cursor at line {line}):");
					sb.AppendLine($"```{ext}");
					sb.AppendLine(code);
					sb.AppendLine("```");
				}
			}
			catch { }

			return sb.ToString();
		}

		private string GetDiagnostics(DTE2 dte)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var sb = new StringBuilder();
			int count = 0;

			try
			{
				ErrorItems items = dte.ToolWindows.ErrorList.ErrorItems;
				for (int i = 1; i <= items.Count && count < 100; i++)
				{
					ErrorItem item = items.Item(i);
					string file = string.IsNullOrEmpty(item.FileName) ? "" : item.FileName;
					string project = string.IsNullOrEmpty(item.Project) ? "" : $"[{item.Project}] ";
					sb.AppendLine($"{project}{file}:{item.Line} - {item.Description}");
					count++;
				}
			}
			catch (Exception ex)
			{
				sb.AppendLine($"Error accessing Error List: {ex.Message}");
			}

			if (count == 0)
				sb.AppendLine("No errors or warnings.");

			return sb.ToString();
		}

		private string GetOpenFiles(DTE2 dte)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var sb = new StringBuilder();

			string solutionDir = null;
			if (dte.Solution != null && !string.IsNullOrEmpty(dte.Solution.FullName))
				solutionDir = Path.GetDirectoryName(dte.Solution.FullName);

			string activeFile = dte.ActiveDocument?.FullName;

			foreach (Document doc in dte.Documents)
			{
				string path = doc.FullName;
				if (!string.IsNullOrEmpty(solutionDir) && path.StartsWith(solutionDir, StringComparison.OrdinalIgnoreCase))
					path = path.Substring(solutionDir.Length).TrimStart('\\', '/');

				string marker = doc.FullName == activeFile ? " (active)" : "";
				sb.AppendLine($"{path}{marker}");
			}

			if (sb.Length == 0)
				sb.AppendLine("No open documents.");

			return sb.ToString();
		}

		private void SendJsonResponse(HttpListenerContext context, object data)
		{
			string jsonStr = json.Serialize(data);
			byte[] buffer = Encoding.UTF8.GetBytes(jsonStr);
			context.Response.ContentType = "application/json";
			context.Response.ContentLength64 = buffer.Length;
			context.Response.OutputStream.Write(buffer, 0, buffer.Length);
			context.Response.Close();
		}

		public void Dispose()
		{
			running = false;
			try { listener?.Stop(); }
			catch { }
			try { listener?.Close(); }
			catch { }
		}
	}
}
