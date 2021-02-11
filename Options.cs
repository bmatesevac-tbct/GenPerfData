using CommandLine;

namespace ImportGenerator
{
   public class Options
   {
      [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
      public bool Verbose { get; set; }

      [Option('d', "devices", Required = true, HelpText = "Specify number of devices.")]
      public int NumDevices { get; set; }

      [Option('f', "facilities",
         Required = true,
         HelpText = "Specify number of facilities.")]
      public int NumFacilities { get; set; }

      [Option('s', "serial", Required = false, HelpText = "Specify the serial number prefix.")]
      public string SerialPrefix { get; set; } = "1X";

      [Option('r', "duration", Required = false, HelpText = "Specify duration in hours (>= 0.1).")]
      public double Duration { get; set; } = 1.1;
   }
}