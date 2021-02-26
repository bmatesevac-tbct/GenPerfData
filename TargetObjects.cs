using RandomNameGeneratorLibrary;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ImportGenerator
{

   public class Facility
   {
      public string Id;
      public string Name;
      public string City;
      public string State;
      public string ZoneId;
      public bool Active;
      public FacilitySchedule Schedule;
      public int OperatingHoursDlogThrottleSpeed;
      public int NonOperatingHoursDlogThrottleSpeed;
   }

   public class FacilityScheduledEvent
   {
      public string FacilityId;
      public int DayOfWeek;
      public string DataTransmissionStart;
      public double DurationHours;
   }

   public class FacilitySchedule : List<FacilityScheduledEvent>
   {
   }

   public class Device
   {
      public string Name;
      public string SerialNumber;
      public string FacilityId;
      public bool Active;
   }

   public class TargetObjects
   {
      public readonly List<Facility> Facilities = new List<Facility>();
      public readonly List<Device> Devices = new List<Device>();
      private readonly Options _options;


      public TargetObjects(Options options)
      {
         _options = options;

         var states = SourceData.States;
         var timeZones = SourceData.TimeZones;

         var placeGenerator = new PlaceNameGenerator();


         int timeZoneIndex = 0;
         int stateIndex = 0;

         for (int nFacility = 1; nFacility <= options.NumFacilities; ++nFacility)
         {
            var timeZone = timeZones[timeZoneIndex];
            timeZoneIndex = ++timeZoneIndex % timeZones.Length;

            var state = states[stateIndex];
            stateIndex = ++stateIndex % states.Length;

            var city = placeGenerator.GenerateRandomPlaceName();

            var facility = new Facility()
            {
               Id = String.Format("Site{0:000}", nFacility),
               Name = String.Format("Facility{0:000}", nFacility),
               City = city,
               State = state,
               ZoneId = timeZone,
               Active = true,
               OperatingHoursDlogThrottleSpeed = 1000,
               NonOperatingHoursDlogThrottleSpeed = 10000
            };
            Facilities.Add(facility);

            // create scheduled events
            var schedule = new FacilitySchedule();
            facility.Schedule = schedule;
            int start = 0;
            for (int day = 1; day <= 7; ++day)
            {
               var scheduledEvent = new FacilityScheduledEvent()
               {

                  FacilityId = facility.Id,
                  DayOfWeek = day,
                  DataTransmissionStart = String.Format("00:{0:00}", start++),
                  DurationHours = options.Duration
               };
               schedule.Add(scheduledEvent);
            }
         }

         // create devices
         // alternate putting devices into the facilities

         var facilityIndex = 0;
         for (int nDevice = 1; nDevice <= options.NumDevices; ++nDevice)
         {
            var facility = Facilities[facilityIndex];
            facilityIndex = ++facilityIndex % Facilities.Count();
            var device = new Device()
            {
               Name = String.Format("Device{0:000}", nDevice),
               SerialNumber = String.Format("{0}{1:00000}", options.SerialPrefix, nDevice),
               FacilityId = facility.Id,
               Active = true,
            };
            Devices.Add(device);
         }
      }

   }




}
