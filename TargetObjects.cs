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

      private string[] _states = SourceData.States;
      private string[] _timeZones = SourceData.TimeZones;
      private PlaceNameGenerator _placeGenerator = new PlaceNameGenerator();

      private int _timeZoneIndex = 0;
      private int _stateIndex = 0;
      private int _nFacilities = 0;
      private int _nDevices = 0;


      public TargetObjects(Options options)
      {
         _options = options;

         var facilityGroupSpecs = _options.FacilityGroupSpecifiers;
         if (facilityGroupSpecs == null)
         {
            for (int nFacility = 0; nFacility < options.NumFacilities; ++nFacility)
            {
               var facility = CreateFacility();
               Facilities.Add(facility);
               facility.Schedule = CreateFacilitySchedule(facility);
            }

            // create devices
            // alternate putting devices into the facilities
            var facilityIndex = 0;
            for (int nDevice = 0; nDevice < options.NumDevices; ++nDevice)
            {
               var facility = Facilities[facilityIndex];
               facilityIndex = ++facilityIndex % Facilities.Count();
               Devices.Add(CreateDevice(facility));
            }
         }

         else
         {
            // distribute based on the facilityGroupSpecs
            foreach (var facilityGroupSpec in facilityGroupSpecs)
            {
               for (int nFacility = 0; nFacility < facilityGroupSpec.NumFacilities; ++nFacility)
               {
                  var facility = CreateFacility();
                  Facilities.Add(facility);
                  facility.Schedule = CreateFacilitySchedule(facility);

                  // add devices to the facility
                  for (int nDevice = 0; nDevice < facilityGroupSpec.NumDevices; ++nDevice)
                  {
                     Devices.Add(CreateDevice(facility));
                  }
               }
            }
         }
      }

      private Facility CreateFacility()
      {
         int nFacility = ++_nFacilities;
         var timeZone = _timeZones[_timeZoneIndex];
         _timeZoneIndex = ++_timeZoneIndex % _timeZones.Length;

         var state = _states[_stateIndex];
         _stateIndex = ++_stateIndex % _states.Length;

         var city = _placeGenerator.GenerateRandomPlaceName();

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

         return facility;
      }

      private FacilitySchedule CreateFacilitySchedule(Facility facility)
      {
         var schedule = new FacilitySchedule();
         int start = 0;
         for (int day = 1; day <= 7; ++day)
         {
            var scheduledEvent = new FacilityScheduledEvent()
            {

               FacilityId = facility.Id,
               DayOfWeek = day,
               DataTransmissionStart = String.Format("00:{0:00}", start++),
               DurationHours = _options.Duration
            };
            schedule.Add(scheduledEvent);
         }
         return schedule;
      }

      private Device CreateDevice(Facility facility)
      {
         int nDevice = ++_nDevices;
         var device = new Device()
         {
            Name = String.Format("Device{0:000}", nDevice),
            SerialNumber = String.Format("{0}{1:00000}", _options.SerialPrefix, nDevice),
            FacilityId = facility.Id,
            Active = true,
         };
         return device;
      }

   }
}



