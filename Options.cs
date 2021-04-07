using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
         var group = groups[0];
         Assert.Equal(1, group.NumFacilities);
         Assert.Equal(2, group.NumDevices);
         group = groups[1];
         Assert.Equal(2, group.NumFacilities);
         Assert.Equal(3, group.NumDevices);
         group = groups[2];
         Assert.Equal(3, group.NumFacilities);
         Assert.Equal(4, group.NumDevices);
      }

      [Fact]
      public void ExtractFacilityGroups_ValidOptions_ProperlyExtracted()
      {
         // test extractions
         var options = new Options();
         var groups = options.ExtractFacilityGroups("1@2,2@3,3@4");
         EnsureFacilityGroupSpecifiers(groups);
         groups = options.ExtractFacilityGroups(" 1 @ 2 ,2@3,3@4 ");
         EnsureFacilityGroupSpecifiers(groups);
      }

      [Fact]
      public void ExtractFacilityGroups_InvalidOptions_SolicitsException()
      {
         // test extractions
         var options = new Options();
         var exception = Assert.Throws<Exception>(() => options.ExtractFacilityGroups("1z2,2@3"));
         Assert.Equal(@"Invalid grouping '1z2' in '1z2,2@3'", exception.Message);

         exception = Assert.Throws<Exception>(() => options.ExtractFacilityGroups("1@2,2z3"));
         Assert.Equal(@"Invalid grouping '2z3' in '1@2,2z3'", exception.Message);

         exception = Assert.Throws<Exception>(() => options.ExtractFacilityGroups("1@2,2@A"));
         Assert.Equal(@"Invalid grouping '2@A' in '1@2,2@A'", exception.Message);
      }

      [Fact]
      public void Options_InvalidOptions_SolicitsException()
      {
         // no arguments
         var anyErrors = TestCommandLineInput(String.Empty);
         Assert.NotEmpty(anyErrors);
         Assert.Equal("\nNumber of facilities must be greater than 0.\nNumber of devices must be greater than 0.", anyErrors);

         // missing facilities param
         anyErrors = TestCommandLineInput("-d 2");
         Assert.NotEmpty(anyErrors);
         Assert.Equal("\nNumber of facilities must be greater than 0.", anyErrors);

         // missing device param
         anyErrors = TestCommandLineInput("-f 2");
         Assert.NotEmpty(anyErrors);
         Assert.Equal("\nNumber of devices must be greater than 0.", anyErrors);

         // -g mutually exclusive with -f
         anyErrors = TestCommandLineInput("-f 2 -g 1@2");
         Assert.NotEmpty(anyErrors);
         Assert.Equal("\nOptions -f and -g and mutually exclusive.", anyErrors);

         // -g mutually exclusive with -d
         anyErrors = TestCommandLineInput("-d 2 -g 1@2");
         Assert.NotEmpty(anyErrors);
         Assert.Equal("\nOptions -d and -g and mutually exclusive.", anyErrors);
      }

      [Fact]
      public void Options_ValidOptions_Accepted()
      {
         // standard -f -d syntax
         var anyErrors = TestCommandLineInput("-f 2 -d 2");
         Assert.Empty(anyErrors);

         // group syntax
         anyErrors = TestCommandLineInput("-g 1@2");
         Assert.Empty(anyErrors);

         // multi-group syntax
         anyErrors = TestCommandLineInput("-g 1@2,3@4");
         Assert.Empty(anyErrors);
      }

      private string TestCommandLineInput(string input)
      {
         var help = new StringWriter();
         var parser = new Parser(config =>
         {
            config.HelpWriter = help;
            config.MaximumDisplayWidth = 80;
         });
         var args = input.Split(' ');
         var options = new Options();
         parser.ParseArguments<Options>(args)
            .WithParsed(o => options = o);
         var anyErrors = Options.Validate(options);
         if (anyErrors != String.Empty)
            help.Write(anyErrors);
         return anyErrors;
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

      public List<FacilityGroupSpecifier> ExtractFacilityGroups(string input)
      {
         var facilityGroups = new List<FacilityGroupSpecifier>();
         var pattern = @"\s*(\d+)\s*@\s*(\d+)\s*";

         var groupings = input.Split(',');
         foreach (var grouping in groupings)
         {
            // get rid of all spaces
            var matches = Regex.Matches(grouping, pattern);

            if (matches.Count != 1)
               throw new Exception($"Invalid grouping '{grouping}' in '{input}'");

            var match = matches[0];
            if (match.Groups.Count != 3)
               throw new Exception($"Invalid grouping '{grouping}' in '{input}'");

            facilityGroups.Add(new FacilityGroupSpecifier()
            {
               NumFacilities = int.Parse(match.Groups[1].Value),
               NumDevices = int.Parse(match.Groups[2].Value)
            });
         }
         return facilityGroups;
      }

      [Option('s', "serial", Required = false, HelpText = "Specify the serial number prefix.")]
      public string SerialPrefix { get; set; } = "1X";

      [Option('r', "duration", Required = false, HelpText = "Specify duration in hours (>= 0.25).")]
      public double Duration { get; set; } = 1.0;

      public static string Validate(Options options)
      {
         var errors = String.Empty;

         if (options.FacilityGroupSpecifiers == null)
         {
            if (options.NumFacilities < 1)
            {
               errors += "\nNumber of facilities must be greater than 0.";
            }
            if (options.NumDevices < 1)
            {
               errors += "\nNumber of devices must be greater than 0.";
            }
         }
         else
         {
            if (options.NumFacilities != 0)
               errors += "\nOptions -f and -g and mutually exclusive.";
            if (options.NumDevices != 0)
               errors += "\nOptions -d and -g and mutually exclusive.";
         }

         return errors;
      }
   }
}