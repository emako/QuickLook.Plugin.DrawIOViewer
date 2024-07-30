// Copyright © 2024 ema
//
// This file is part of QuickLook program.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using Microsoft.Win32;
using QuickLook.Common.Plugin;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace QuickLook.Plugin.DrawIOViewer;

public class Plugin : IViewer
{
    private static readonly HashSet<string> hashSet =
    [
        ".drawio",
    ];

    private static readonly HashSet<string> WellKnownImageExtensions = hashSet;

    private ImagePanel? _ip;
    private string _drawio = null!;

    public int Priority => 0;

    public void Init()
    {
    }

    public bool CanHandle(string path)
    {
        return WellKnownImageExtensions.Contains(Path.GetExtension(path.ToLower()));
    }

    public void Prepare(string path, ContextObject context)
    {
        context.PreferredSize = new Size { Width = 800, Height = 600 };
    }

    public void View(string path, ContextObject context)
    {
        _ip = new ImagePanel
        {
            ContextObject = context,
        };

        _ = Task.Run(() =>
        {
            byte[] imageData = ViewImage(path);
            BitmapImage bitmap = new();
            using MemoryStream stream = new(imageData);
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.StreamSource = stream;
            bitmap.EndInit();
            bitmap.Freeze();

            _ip.Dispatcher.Invoke(() => _ip.Source = bitmap);

            context.IsBusy = false;
        });

        context.ViewerContent = _ip;
        context.Title = $"{Path.GetFileName(path)}";
    }

    public void Cleanup()
    {
        GC.SuppressFinalize(this);

        _ip = null;
    }

    public byte[] ViewImage(string path)
    {
        _drawio ??= FindDrawIO();

        if (!File.Exists(_drawio))
        {
            _drawio = FindDrawIO();
        }

        string png = Path.GetFullPath(@".\QuickLook.Plugin\QuickLook.Plugin.DrawIOViewer\diagram.png");

        if (File.Exists(png))
        {
            File.Delete(png);
        }

        using Process process = new()
        {
            StartInfo = new ProcessStartInfo()
            {
                FileName = _drawio,
                Arguments = $"--export --format png --output \"{png}\" --scale 1.4 \"{path}\"",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            }
        };

        process.Start();
        process.WaitForExit();

        if (File.Exists(png))
        {
            byte[] bytes = File.ReadAllBytes(png);
            File.Delete(png);
            return bytes;
        }

        return null!;
    }

    private static string FindDrawIO()
    {
        string? uninstallInfo = (GetUninstallInfo(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "draw.io")
                             ?? GetUninstallInfo(@"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "draw.io"))
                             ?? throw new ApplicationException("draw.io is not installed, register not found.");

        string[] parsedArgs = ParseArguments(uninstallInfo);

        if (parsedArgs.Length <= 0)
        {
            throw new ApplicationException("draw.io is not installed, UninstallString is empty.");
        }

        FileInfo uninst = new(parsedArgs[0].Trim('"'));
        string drawio = Path.Combine(uninst.DirectoryName, "draw.io.exe");

        if (!File.Exists(drawio))
        {
            throw new ApplicationException("draw.io is not installed, file draw.io.exe not found.");
        }

        return drawio;
    }

    private static string? GetUninstallInfo(string keyPath, string displayName)
    {
        using RegistryKey key = Registry.LocalMachine.OpenSubKey(keyPath);

        if (key != null)
        {
            foreach (string subkeyName in key.GetSubKeyNames())
            {
                using RegistryKey subkey = key.OpenSubKey(subkeyName);

                if (subkey != null)
                {
                    if (subkey.GetValue("DisplayName") is string name && name.Contains(displayName))
                    {
                        string? uninstallString = subkey.GetValue("UninstallString") as string;

                        if (!string.IsNullOrEmpty(uninstallString))
                        {
                            return uninstallString;
                        }
                    }
                }
            }
        }
        return null;
    }

    private static string[] ParseArguments(string commandLine)
    {
        List<string> args = [];
        string currentArg = string.Empty;
        bool inQuotes = false;

        for (int i = 0; i < commandLine.Length; i++)
        {
            char c = commandLine[i];

            if (c == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (c == ' ' && !inQuotes)
            {
                if (currentArg != string.Empty)
                {
                    args.Add(currentArg);
                    currentArg = string.Empty;
                }
            }
            else
            {
                currentArg += c;
            }
        }

        if (currentArg != string.Empty)
        {
            args.Add(currentArg);
        }

        return [.. args];
    }
}
