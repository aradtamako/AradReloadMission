using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.Linq;
using WindowsInput;
using System.Threading;
using System.Windows.Forms;
using WindowsInput.Native;
using System.Collections.Generic;

namespace AradReloadMission
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        static InputSimulator Simulator { get; } = new InputSimulator();

        static double GetAbsoluteX(int x) => x * 65535 / Screen.PrimaryScreen.Bounds.Width;
        static double GetAbsoluteY(int y) => y * 65535 / Screen.PrimaryScreen.Bounds.Height;

        static int AradClientWidth = 1440;
        static int AradClientHeight = 900;

        static void MouseLeftClick()
        {
            Simulator.Mouse.LeftButtonDown();
            Thread.Sleep(50);
            Simulator.Mouse.LeftButtonUp();
        }

        static void MouseMoveAndLeftClick(int x, int y)
        {
            Simulator.Mouse.MoveMouseTo(GetAbsoluteX(x), GetAbsoluteY(y));
            MouseLeftClick();
        }

        static void PressKeyboard(VirtualKeyCode virtualKeyCode)
        {
            Simulator.Keyboard.KeyDown(virtualKeyCode);
            Thread.Sleep(50);
            Simulator.Keyboard.KeyUp(virtualKeyCode);
        }

        /// <summary>
        /// 指定した範囲のスクリーンショットを撮る
        /// </summary>
        static Bitmap TakeScreenShot(int x1, int y1, int x2, int y2)
        {
            var width = x2 - x1;
            var height = y2 - y1;
            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(x1, y1, 0, 0, bmp.Size);
            }
            return bmp;
        }

        static Bitmap ConvertTo32bppRgb(Bitmap bmp)
        {
            if (bmp.PixelFormat == PixelFormat.Format32bppRgb) { return bmp; }

            var bmp2 = new Bitmap(bmp.Width, bmp.Height, PixelFormat.Format32bppRgb);

            using (var g = Graphics.FromImage(bmp2))
            {
                g.PageUnit = GraphicsUnit.Pixel;
                g.DrawImageUnscaled(bmp, 0, 0);
            };

            return bmp2;
        }

        /// <summary>
        /// テンプレートマッチを行う
        /// </summary>
        static System.Drawing.Point MatchTemplate(string templateFileName, Bitmap screenBmp)
        {
            try
            {
                using (var templateBmp = new Bitmap(templateFileName))
                using (var screen = BitmapConverter.ToMat(ConvertTo32bppRgb(screenBmp)))
                using (var template = BitmapConverter.ToMat(ConvertTo32bppRgb(templateBmp)))
                using (var result = new Mat())
                {
                    Cv2.MatchTemplate(screen, template, result, TemplateMatchModes.CCoeffNormed);
                    Cv2.Threshold(result, result, 0.8, 1.0, ThresholdTypes.Binary);
                    Cv2.MinMaxLoc(result, out OpenCvSharp.Point minPoint, out OpenCvSharp.Point maxPoint);
                    return new System.Drawing.Point(maxPoint.X, maxPoint.Y);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                return new System.Drawing.Point(0, 0);
            }
        }

        static void Initialize()
        {
            var process = Process.GetProcessesByName("ARAD").FirstOrDefault();
            if (process == null)
            {
                throw new Exception("Process not found");
            }

            var hWndDNF = process.MainWindowHandle;
            if (hWndDNF != IntPtr.Zero)
            {
                GetWindowRect(hWndDNF, out var rect);
                if (rect.Width != AradClientWidth || rect.Height != AradClientHeight)
                {
                    Console.WriteLine($"アラド戦記の解像度を{AradClientWidth}x{AradClientHeight}に設定して下さい。");
                    return;
                }
                MoveWindow(hWndDNF, 0, 0, rect.Width, rect.Height, true);
            }
        }

        static void Main()
        {
            Initialize();

            Console.WriteLine("何かキーを押して下さい");
            Console.ReadKey();

            var offset = 100;

            for (var i = 0; i < 3; i++)
            {
                var points = new List<System.Drawing.Point>();

                while (true)
                {
                    using var bmp = TakeScreenShot(207, 307 + (offset * i), 456, 361 + (offset * i));
                    foreach (var file in Directory.EnumerateFiles("templates"))
                    {
                        var point = MatchTemplate(file, bmp);
                        points.Add(point);
                        Console.WriteLine(file);
                        Console.WriteLine(point);
                    }

                    if (points.Where(x => x.X != 0 || x.Y != 0).Any())
                    {
                        points.Clear();
                        break;
                    }
                    else
                    {
                        MouseMoveAndLeftClick(444, 373 + (offset * i));
                        PressKeyboard(VirtualKeyCode.SPACE);
                    }

                    // ミッション文字列がフェードインで表示されるのでディレイを入れる
                    // Thread.Sleep(200);
                }
            }
        }
    }
}
