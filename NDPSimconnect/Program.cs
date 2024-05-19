using System;
using System.Collections.Generic;
using System.Threading;
using System.IO;
using System.Net.Http;
using System.Globalization;
using FSUIPC;
using Microsoft.FlightSimulator.SimConnect;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.Diagnostics;

namespace NDPSimconnect
{
    class Program
    {
        // FSUIPC variables, used when running in FSUIPC mode
        private Offset<FsLongitude> playerLon = new Offset<FsLongitude>("NDP", 0x0568, 8);
        private Offset<FsLatitude> playerLat = new Offset<FsLatitude>("NDP", 0x0560, 8);
        private Offset<uint> playerHdg = new Offset<uint>("NDP", 0x0580);
        private Offset<uint> playerAlt = new Offset<uint>("NDP", 0x0020);

        // SimConnect variables, used when running in SimConnect mode
        private SimConnect simconnect = null;
        private enum DEFINITIONS
        {
            SimConnectData,
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        private struct SimConnectData
        {
            public double latitude;
            public double longitude;
            public double heading;
            public double altitude;
        };

        private string ndp_charts_settings = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), ".ndp-chartcloud\\ndp-settings.json");
        private NDPsettings settings = null;
        private HttpClient client = new HttpClient();

        private Timer timer = null;

        public class NDPsettings
        {
            public string sessionId = "";
        }

        static void Main(string[] args)
        {
            Program prog = new Program();
            prog.Run();
        }

        private void Run()
        {
            // find the open NDP session, without this we can't do anything
            try
            {
                findOpenNDPSession();
            } catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return;
            }

            bool connected = false;

            // try connecting to the simulator via SimConnect, if that fails try FSUIPC/XPUIPC
            do {
                Console.WriteLine("Trying to connect to Simulator...");
                try
                {
                    openSimConnect();
                    connected = true;
                } catch (Exception simconnectEx) {
                    this.simconnect = null;     // unset the simconnect object so we don't try to use it later

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Could not open SimConnect connection!");
                    Console.WriteLine(simconnectEx.Message);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Trying to connect to FSUIPC/XPUIPC instead...");
                    Console.ResetColor();

                    try
                    {
                        FSUIPCConnection.Open();
                        connected = true;
                    } catch (Exception fsuipcEx)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Could not open FSUIPC/XPUIPC connection!");
                        Console.WriteLine(fsuipcEx.Message);
                        Console.ResetColor();

                        return;
                    }
                }

                // wait for a second before trying again
                Thread.Sleep(1000);
            } while (!connected);

            // if using SimConnect, run the message loop
            if (this.simconnect != null)
            {
                while (true)
                {
                    this.simconnect.ReceiveMessage();
                    Thread.Sleep(100);
                }
            }
            else if (FSUIPCConnection.IsOpen)
            {
                // if using FSUIPC, run the other message loop
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Successfully connected to simulator (FSUIPC)");
                Console.ResetColor();

                timer = new Timer(fsuipcCallback, null, 0, 1000);

                while (true)
                {
                    Thread.Sleep(100);
                }
            }
        }

        /// <summary>
        /// Find the current NDP session ID from the settings file
        /// </summary>
        /// <exception cref="Exception"></exception>
        private void findOpenNDPSession()
        {
            if (File.Exists(ndp_charts_settings))
            {
                settings = JsonConvert.DeserializeObject<NDPsettings>(File.ReadAllText(ndp_charts_settings).Trim());
                if (settings != null && settings.sessionId.Length > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Found existing NDP session: {settings.sessionId}");
                    Console.ResetColor();
                } else
                {
                    throw new Exception("No active NDP session!");
                }
            } else
            {
                throw new Exception("Could not find NDP settings file!");
            }
        }


        private void openSimConnect()
        {
            this.simconnect = new SimConnect("NDP SimConnect", IntPtr.Zero, 0, null, 0);
            this.simconnect.OnRecvOpen += simconnect_OnRecvOpen;

            this.simconnect.AddToDataDefinition(DEFINITIONS.SimConnectData, "Plane Latitude", "degrees", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
            this.simconnect.AddToDataDefinition(DEFINITIONS.SimConnectData, "Plane Longitude", "degrees", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
            this.simconnect.AddToDataDefinition(DEFINITIONS.SimConnectData, "Plane Heading Degrees True", "degrees", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
            this.simconnect.AddToDataDefinition(DEFINITIONS.SimConnectData, "Plane Altitude", "feet", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
            
            this.simconnect.RegisterDataDefineStruct<SimConnectData>(DEFINITIONS.SimConnectData);
            this.simconnect.RequestDataOnSimObject(DEFINITIONS.SimConnectData, DEFINITIONS.SimConnectData, 0, SIMCONNECT_PERIOD.SECOND, SIMCONNECT_DATA_REQUEST_FLAG.CHANGED, 0, 1, 0);
            this.simconnect.OnRecvSimobjectData += simconnect_OnRecvSimobjectData;

        }

        private void simconnect_OnRecvOpen(SimConnect sender, SIMCONNECT_RECV_OPEN data)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Successfully connected to simulator (SimConnect)");
            Console.ResetColor();
        }

        private void simconnect_OnRecvSimobjectData(SimConnect sender, SIMCONNECT_RECV_SIMOBJECT_DATA data)
        {
            SimConnectData simData = (SimConnectData)data.dwData[0];

            sendData(simData.latitude, simData.longitude, simData.heading, simData.altitude);
        }

        /// <summary>
        /// Timer callback to update the location, runs every 500ms
        /// </summary>
        /// <param name="o"></param>
        private void fsuipcCallback(Object o)
        {
            FSUIPCConnection.Process("NDP");

            double lat = playerLat.Value.DecimalDegrees;
            double lon = playerLon.Value.DecimalDegrees;
            double hdg = playerHdg.Value * 360.0 / (65535.0 * 65535.0);
            double alt = playerAlt.Value / 0.3048 / 256.0;

            sendData(lat, lon, hdg, alt);
        }

        private void sendData(double lat, double lon, double hdg, double alt)
        {
            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";

            Dictionary<string, string> formData = new Dictionary<string, string>
            {
                {
                    "sessionId",
                    settings.sessionId
                },
                {
                    "latitude",
                    lat.ToString("0.######", nfi)
                },
                {
                    "longitude",
                    lon.ToString("0.######", nfi)
                },
                {
                    "heading",
                    hdg.ToString("0.##", nfi)
                },
                {
                    "altitude",
                    alt.ToString("0.##", nfi)
                }
            };

            try
            {
                FormUrlEncodedContent content = new FormUrlEncodedContent(formData);
                Console.WriteLine($"Lat: {lat}, Lon: {lon}, Hdg: {hdg}, Alt: {alt}");
                HttpResponseMessage response = client.PostAsync("https://navdatapro.aerosoft.com/api/v3/location", content).Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating location! {ex}");
            }
        }   
    }
}
