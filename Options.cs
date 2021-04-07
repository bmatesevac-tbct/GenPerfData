using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using CommandLine;
using Xunit;
using Xunit.Abstractions;


namespace ImportGenerator
{

   public class OptionsUnitTests
   {

      private void EnsureFacilityGroupSpecifiers(List<Options.FacilityGroupSpecifier> groups)
      {
         Assert.NotNull(groups);
         Assert.Equal(3, groups.Count);
      }

      [Fact]
      public void ExtractFacilityGroups_ExtractsValidGroups()
      {
         // test extractions
         var options = new Options();
         var groups = options.ExtractFacilityGroups("1@2,2@3,3@4");
         EnsureFacilityGroupSpecifiers(groups);
         groups = options.ExtractFacilityGroups(" 1 @ 2 ,2@3,3@4 ");
         EnsureFacilityGroupSpecifiers(groups);
      }

   }



   public class Options
   {
      public class FacilityGroupSpecifier
      {
         public int NumFacilities;
         public int NumDevices;
      }

      [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
      public bool Verbose { get; set; }

      [Option('o', "output",
         Required = false,
         HelpText = "Specify output file name.")]
      public string OutputFileName { get; set; }

      [Option('d', "devices", 
         Required = false, 
         HelpText = "Specify number of devices.")]
      public int NumDevices { get; set; } = 0;

      [Option('f', "facilities",
         Required = false,
         HelpText = "Specify number of facilities.")]
      public int NumFacilities { get; set; } = 0;


      public List<FacilityGroupSpecifier> FacilityGroupSpecifiers { get; private set; }

      [Option('g', "groups",
         Required = false,
         HelpText = "Specify facility/device count groupings <ff@dd,ff@dd,...> ex: 1@20,10@30 will produce 1 facility with 20 device and 10 facilities with 30 devices.")]

      public string FacilityGroup
      {
         set
         {
            FacilityGroupSpecifiers = ExtractFacilityGroups(value);
            if ((FacilityGroupSpecifiers == null) || (FacilityGroupSpecifiers.Count < 1))
               throw new Exception($"Invalid group specifier {value}");
         }
      }

      public List<FacilityGroupSpecifier> ExtractFacilityGroups(string text)
      {
         var facilityGroups = new List<FacilityGroupSpecifier>();
         var pattern = @"\s*(\d+)\s*@\s*(\d+)\s*";

         var groupings = text.Split(',');
         foreach (var grouping in groupings)
         {
            // get rid of all spaces
            var matches = Regex.Matches(grouping, pattern);

            if (matches.Count != 1)
               throw new Exception($"Invalid grouping '{grouping}' in '{text}'");

            var match = matches[0];
            if (match.Groups.Count != 3)
               throw new Exception($"Invalid grouping '{grouping}' in '{text}'");

            facilityGroups.Add(new FacilityGroupSpecifier()
            {
               NumDevices = int.Parse(match.Groups[1].Value),
               NumFacilities = int.Parse(match.Groups[2].Value)
            });
         }
         return facilityGroups;
      }

      [Option('s', "serial", Required = false, HelpText = "Specify the serial number prefix.")]
      public string SerialPrefix { get; set; } = "1X";

      [Option('r', "duration", Required = false, HelpText = "Specify duration in hours (>= 0.25).")]
      public double Duration { get; set; } = 1.0;
   }
}