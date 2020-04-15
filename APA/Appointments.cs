using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace APA
{
    public class Appointments : List<Appointment>
    {        
        public Appointments(string mail, ExchangeService service)
        {
            try
            {
                CalendarView calView = new CalendarView(
                    DateTime.Now,  // von 
                    Global.LetzterTagDesSchuljahres) // bis
                {
                    PropertySet = new PropertySet(
                        BasePropertySet.IdOnly,
                        AppointmentSchema.Subject,
                        AppointmentSchema.Start,
                        AppointmentSchema.IsRecurring,
                        AppointmentSchema.AppointmentType,
                        AppointmentSchema.Categories)
                };

                FindItemsResults<Appointment> findResults = service.FindAppointments(
                    WellKnownFolderName.Calendar,
                    calView);

                List<Appointment> relevanteAppointments = new List<Appointment>();

                Console.WriteLine("Existierende Appointments für " + mail + ":");

                foreach (var appointment in findResults.Items)
                {
                    // Alle relevanten Appointments für diese Zielperson werden in eine Liste geladen, ...

                    appointment.Load();

                    // ... um die Eigenschaften von Appointment und Termin vergleichen zu können.

                    if (appointment.Categories.Contains("ZulassungskonferenzenBC"))
                    {
                        // Wenn Subject oder Body null sind, werden sie durch "" ersetzt.

                        appointment.Subject = appointment.Subject ?? "";

                        appointment.Body = appointment.Body ?? "";

                        this.Add(appointment);
                        Console.WriteLine("[Ist ] " + appointment.Subject.PadRight(30) + appointment.Start + "-" + appointment.End.ToShortTimeString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler beim Lehrer mit der Adresse " + mail + "\n" + ex.ToString());
                throw new Exception("Fehler beim Lehrer mit der Adresse " + mail + "\n" + ex.ToString());
            }
        }

        public Appointments(Excelzeilen excelzeilen, ExchangeService service)
        {
            try
            {
                foreach (var excelzeile in (from e in excelzeilen where e.ADatum >= DateTime.Now select e).ToList())
                {
                    Appointment appointment = new Appointment(service);
                    appointment.Categories.Add("ZulassungskonferenzenBC");
                    appointment.Start = new DateTime(excelzeile.ADatum.Year, excelzeile.ADatum.Month, excelzeile.ADatum.Day, excelzeile.FBeginn.Hour, excelzeile.FBeginn.Minute, excelzeile.FBeginn.Second);
                    appointment.IsAllDayEvent = excelzeile.FBeginn.Hour == 0 ? true : false;
                    appointment.End = (excelzeile.FBeginn.Hour == 0 ? excelzeile.FBeginn.AddDays(1) : new DateTime(excelzeile.ADatum.Year, excelzeile.ADatum.Month, excelzeile.ADatum.Day, excelzeile.GEnde.Hour, excelzeile.GEnde.Minute, excelzeile.GEnde.Second));
                    appointment.IsReminderSet = false;
                    appointment.Location = excelzeile.HRaum != null ? excelzeile.HRaum.Count > 0 ? excelzeile.HRaum[0].RaumApa : "" : "";
                    appointment.Subject = excelzeile.Subject;
                    appointment.Body = excelzeile.DBeschreibung;

                    this.Add(appointment);

                    Console.WriteLine("[Soll] " + appointment.Subject.PadRight(30) + appointment.Start + "-" + appointment.End.ToShortTimeString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
        }

        internal void DeleteAppointments(Appointments appointmentsSoll)
        {
            try
            {
                foreach (var appointmentIst in this)
                {
                    // Wenn das Ist-Appointment nicht in den Soll-Appointments existiert, ... 

                    if (!(from a in appointmentsSoll
                          where a.Subject == appointmentIst.Subject
                          where (a.Start == appointmentIst.Start && a.End == appointmentIst.End) || (a.Start.Date == appointmentIst.Start.Date && a.IsAllDayEvent == appointmentIst.IsAllDayEvent && appointmentIst.IsAllDayEvent == true)
                          where (appointmentIst.Categories.Contains("ZulassungskonferenzenBC"))
                          select a).Any())
                    {
                        // ... wird es gelöscht.

                        appointmentIst.Delete(
                            DeleteMode.HardDelete);

                        Console.WriteLine("[ -  ] " + appointmentIst.Subject.PadRight(30) + appointmentIst.Start + " - " + appointmentIst.End);
                    }
                }

                // Wenn ein Ist-Appointment aus irgendeinem Grund mehrfach angelegt wurde, ...

                var x = (from t in this orderby t.Start, t.End, t.Subject select t).ToList();

                for (int i = 1; i < x.Count(); i++)
                {
                    if (x[i].Subject == x[i - 1].Subject && x[i].Start == x[i - 1].Start && x[i].End == x[i - 1].End && (x[i].Categories.Contains("ZulassungskonferenzenBC") && x[i - 1].Categories.Contains("ZulassungskonferenzenBC")))
                    {
                        // ... wird der Vorgänger gelöscht.

                        Console.WriteLine("[ -  ] " + x[i].Subject.PadRight(30) + x[i].Start + " - " + x[i].End + " gelöscht, da mehrfach existent.");

                        x[i - 1].Delete(
                            DeleteMode.HardDelete);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
        }

        internal void AddAppointments(
            Appointments appointmentsIst,
            Lehrer lehrer,
            ExchangeService service)
        {
            try
            {
                foreach (var appointmentSoll in this)
                {
                    // Wenn das Soll-Appointment nicht in den Ist-Appointments existiert, ... 

                    if (!(from a in appointmentsIst
                          where a.Subject == appointmentSoll.Subject
                          where a.Start == appointmentSoll.Start
                          where a.End == appointmentSoll.End
                          where (appointmentSoll.Categories.Contains("ZulassungskonferenzenBC"))
                          select a).Any())
                    {
                        // ... wird angelegt

                        service.ImpersonatedUserId = new ImpersonatedUserId(
                            ConnectingIdType.SmtpAddress,
                            lehrer.Mail);

                        appointmentSoll.IsReminderSet = false;
                        appointmentSoll.Save();

                        Console.WriteLine("[ +  ] " + appointmentSoll.Subject.PadRight(30) + appointmentSoll.Start + " - " + appointmentSoll.End);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
        }
    }
}