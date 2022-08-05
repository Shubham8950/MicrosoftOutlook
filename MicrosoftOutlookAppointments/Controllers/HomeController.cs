using Microsoft.Exchange.WebServices.Data;
using MicrosoftOutlookAppointments.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MicrosoftOutlookAppointments.Controllers
{
    public class HomeController : Controller
    {

        public ActionResult Index()
        {
            string ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx";
            string userName = "youroutlookemail";
            string password = "youroutlookpassword";

            ExchangeService servicex = new ExchangeService();
            servicex.Url = new Uri(ewsUrl);
            servicex.UseDefaultCredentials = true;
            servicex.Credentials = new WebCredentials(userName, password);
            DateTime startDate = DateTime.Now;
            DateTime endDate = startDate.AddDays(30);
            const int NUM_APPTS = 5;
            // Initialize the calendar folder object with only the folder ID. 
            CalendarFolder calendar = CalendarFolder.Bind(servicex, WellKnownFolderName.Calendar, new PropertySet());
            // Set the start and end time and number of appointments to retrieve.
            CalendarView cView = new CalendarView(startDate, endDate, NUM_APPTS);
            // Limit the properties returned to the appointment's subject, start time, and end time.
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);
            // Retrieve a collection of appointments by using the calendar view.
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);
            Debug.WriteLine("\nThe first " + NUM_APPTS + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
                              " to " + endDate.Date.ToShortDateString() + " are: \n");
            List<clsAppointment> listAppointments = new List<clsAppointment>();
            foreach (Appointment a in appointments)
            {
                /*Here you will get your appointments*/
                Debug.Write("Subject: " + a.Subject.ToString() + " ");
                Debug.Write("Start: " + a.Start.ToString() + " ");
                Debug.Write("End: " + a.End.ToString());
                clsAppointment app = new clsAppointment();
                app.Subject = a.Subject.ToString();
                app.StartDate = a.Start;
                app.EndDate = a.End;
                listAppointments.Add(app);

            }
            return View(listAppointments);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }
        public ActionResult GetAppointments()
        {
           
            return View();
        }
        public ActionResult AddAppointments()
        {
            string ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx";
            string userName = "youroutlookemail";
            string password = "youroutlookpassword";

            ExchangeService servicex = new ExchangeService();
            servicex.Url = new Uri(ewsUrl);
            servicex.UseDefaultCredentials = true;
            servicex.Credentials = new WebCredentials(userName, password);


            Appointment appointment = new Appointment(servicex);
            
            appointment.Subject = "Code2night Event";
            appointment.Body = "Focus on backhand this week.";
            appointment.Start = DateTime.Now.AddDays(1);
            appointment.End = appointment.Start.AddHours(1);
            appointment.Location = "Tennis club";
            appointment.ReminderDueBy = DateTime.Now;


            // Save the appointment to your calendar.
            appointment.Save(SendInvitationsMode.SendToNone);
            return RedirectToAction("Index");
        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}