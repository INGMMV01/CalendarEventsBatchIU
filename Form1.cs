using System;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace CalendarEventsBatchV2
{
    public partial class Form1 : Form
    {

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        //kitos static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        System.Collections.Hashtable htCalendars = new System.Collections.Hashtable();
        CalendarService objServicio = null;
        System.Text.StringBuilder stringbuilder = new System.Text.StringBuilder();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ObtenerCalendarios();
        }

        private void ObtenerCalendarios()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/calendar-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            objServicio = service;

            CalendarListResource.ListRequest requestListCalendar = service.CalendarList.List();

            CalendarList calendars = requestListCalendar.Execute();
            foreach (CalendarListEntry cal in calendars.Items)
            {
                System.Drawing.Color color = System.Drawing.Color.FromName(cal.ColorId);
                this.lbTodos.ForeColor = color;
                this.lbTodos.Items.Add(cal.Summary);
                this.htCalendars.Add(cal.Summary, cal);
            }            
        }

        private void btnSeleccionar_Click(object sender, EventArgs e)
        {
            MoverItemListBox(lbTodos, lbElegidos);
        }

        private void btnDescartar_Click(object sender, EventArgs e)
        {
            MoverItemListBox(lbElegidos, lbTodos);
        }

        private void MoverItemListBox(ListBox origen, ListBox destino)
        {
            System.Collections.ArrayList itemsEliminar = new System.Collections.ArrayList();
            foreach (Object item in origen.SelectedItems)
            {
                destino.Items.Add(item);
                itemsEliminar.Add(item);
            }

            foreach(object obj in itemsEliminar)
            {
                origen.Items.Remove(obj);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            stringbuilder = new System.Text.StringBuilder();
            richTextBox1.Text = String.Empty;

            foreach (string nombreCalendario in lbElegidos.Items)
            {
                CalendarListEntry cal = (CalendarListEntry)htCalendars[nombreCalendario];                

                Events eventos = ObtenerEventos(cal);
                if (eventos.Items != null && eventos.Items.Count > 0)
                {
                    foreach (var eventItem in eventos.Items)
                    {
                        ModificarEvento(cal, eventItem);
                    }
                }
            }

            richTextBox1.Text = stringbuilder.ToString();
        }

        private Events ObtenerEventos(CalendarListEntry cal)
        {
            // Define parameters of request.
            EventsResource.ListRequest request = objServicio.Events.List(cal.Id);
            request.TimeMin = this.dateTimePicker1.Value.Date;
            request.TimeMax = this.dateTimePicker1.Value.Date.AddDays(1);
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 100;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
            
            Events events = request.Execute();
            return events;
        }

        private void ModificarEvento(CalendarListEntry cal, Event eventItem)
        {
            TimeSpan ts = dateTimePicker2.Value - dateTimePicker1.Value;
            int days = ts.Days;

            string fechaHoraAntes = eventItem.Start.DateTime.ToString();

            if (eventItem.Start.DateTime.HasValue)
                eventItem.Start.DateTime = eventItem.Start.DateTime.Value.AddDays(days);
            else
            {
                return;

                //todo: no rula esta asignación
                DateTime dt = DateTime.Parse(eventItem.Start.Date);
                //eventItem.Start.Date = dt.AddDays(days).Date.ToString();
                eventItem.Start.DateTime = dt.AddDays(days);
            }

            if (eventItem.End.DateTime.HasValue)
                eventItem.End.DateTime = eventItem.End.DateTime.Value.AddDays(days);
            else
            {
                return;

                //todo: no rula esta asignación
                DateTime dt = DateTime.Parse(eventItem.End.Date);
                //eventItem.End.Date = dt.AddDays(days).Date.ToString();
                eventItem.End.DateTime = dt.AddDays(days);
            }
            
            EventsResource.PatchRequest patchRequest = objServicio.Events.Patch(eventItem, cal.Id, eventItem.Id);
            Event event2 = patchRequest.Execute();
            
            stringbuilder.AppendLine(String.Format("{0} - {1}: ({2} a {3})", cal.Summary, event2.Summary, fechaHoraAntes, event2.Start.DateTime.ToString()));
        }

    }
}
