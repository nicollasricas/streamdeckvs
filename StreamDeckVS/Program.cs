using System;
using BarRaider.SdTools;

namespace StreamDeckVS
{
    internal class Program
    {
        [STAThread]
        private static void Main(string[] args) => SDWrapper.Run(args);
    }
}