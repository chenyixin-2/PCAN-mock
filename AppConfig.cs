using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Peak.Can.Basic;
using Peak.Can.Basic.BackwardCompatibility;
using TPCANHandle = System.UInt16;

public class AppConfig
{
    public Dictionary<string, TPCANBaudrate> DeviceBaudRates { get; set; } = new Dictionary<string, TPCANBaudrate>();

    private static string configFilePath = "appconfig.json";

    public static AppConfig Load()
    {
        if (File.Exists(configFilePath))
        {
            string json = File.ReadAllText(configFilePath);
            return JsonSerializer.Deserialize<AppConfig>(json);
        }
        return new AppConfig();
    }

    public void Save()
    {
        string json = JsonSerializer.Serialize(this);
        File.WriteAllText(configFilePath, json);
    }
}
