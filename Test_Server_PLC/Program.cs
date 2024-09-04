using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using EasyModbus;

namespace Test_Server_PLC
{
    class Program
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        const uint WM_CLOSE = 0x0010;
        private const byte VK_F10 = 0x79; // Código virtual para F10
        private static bool previousBoolValue;

        static void Main(string[] args)
        {
            string plcIp = "192.168.0.115";  // IP del PLC Micro820
            int port = 502;  // Puerto estándar de Modbus TCP
            string softwarePath = @"C:\Users\juan_\OneDrive\Desktop\HMB_Autofocus\HBM Marking Control Software V4_0_1\HBM.exe";

           
            while (true)
            {
                try
                {
                    // Crear un cliente Modbus usando EasyModbus
                    ModbusClient modbusClient = new ModbusClient(plcIp, port);

                    // Conectar al PLC
                    modbusClient.Connect();

                    int previousValue = -1;
                    bool previosBoolValue = false;

                    while (true)
                    {
                        try
                        {
                            int[] holdingRegisters = modbusClient.ReadHoldingRegisters(0, 1);

                            if (holdingRegisters[0] != previousValue)
                            {
                                previousValue = holdingRegisters[0];
                                Console.WriteLine($"Valor leído del PLC: {holdingRegisters[0]}");

                                if (holdingRegisters[0] == 550)
                                {
                                    LaunchAndPrepareSoftware(softwarePath);
                                }
                            }
                            ushort boolStartAddress = 0;
                            short numCoils = 1 ;
                            bool[] coilsStatus = modbusClient.ReadCoils(boolStartAddress, numCoils);

                            if (coilsStatus[0] != previousBoolValue) 
                            {
                                previousBoolValue = coilsStatus[0];
                                Console.WriteLine($"Direccion BOOl {boolStartAddress}: {previousBoolValue} ");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error durante la lectura de datos: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error de conexión: {ex.Message}");
                    Thread.Sleep(5000); // Espera 5 segundos antes de intentar reconectar
                }
            }
        }

        static void LaunchAndPrepareSoftware(string softwarePath)
        {
            Process.Start(softwarePath);
            Thread.Sleep(2000); // Espera para que aparezcan las ventanas emergentes

            ClosePopup("HL"); // Intenta cerrar la primera ventana emergente
            Thread.Sleep(1000); // Espera un poco antes de cerrar la siguiente
            ClosePopup("HL"); // Intenta cerrar la segunda ventana emergente
            Thread.Sleep(1000); // Espera un poco antes de cerrar la siguiente
            ClosePopup("HL"); // Intenta cerrar la tercera ventana emergente

            // Espera para asegurar que el software esté completamente cargado
            Thread.Sleep(3000);

            // Ahora presiona F10
            SendF10Key();
        }

        static void ClosePopup(string windowTitle)
        {
            IntPtr hWnd = FindWindow(null, windowTitle);
            if (hWnd != IntPtr.Zero)
            {
                PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }

        static void SendF10Key()
        {
            const int KEYEVENTF_EXTENDEDKEY = 0x1;
            const int KEYEVENTF_KEYUP = 0x2;

            // Presionar tecla F10
            keybd_event(VK_F10, 0, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero);

            // Liberar tecla F10
            keybd_event(VK_F10, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    }
}
