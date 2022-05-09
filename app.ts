import { calendar_v3, google } from "googleapis";
import { TOKEN } from "./enviroment/variables";
import XLSX from "xlsx";
import { TypeDate } from "./interfaces/dates";

export const getGoogleCalendars = (token: string) => {
  try {
    const oAuth2Client = new google.auth.OAuth2();
    oAuth2Client.setCredentials({ access_token: token });

    const calendar = google.calendar({ version: "v3", auth: oAuth2Client });
    console.log("Results");
    calendar.calendarList.list({}, (err, result) => {
      //console.log(err);
      //console.log(result);
      if (err) {
        console.log("Error al obtener los calendarios de Google calendar");
        return;
      }

      if (!result) {
        console.log("No hubo datos encontrados en Google Calendar");
        return;
      }

      const allCalendars = result.data.items;
      //console.log(allCalendars);
      convertToExcel(allCalendars);
    });
  } catch (e) {
    console.log("error");
  }
};

const convertToExcel = (
  calendars: calendar_v3.Schema$CalendarListEntry[] | undefined
) => {
  if (!calendars) {
    console.log("No hay calnedario en Google ");
    return;
  }
  const calendarsExcel: TypeDate[] = calendars.map((calendar) => {
    if (!calendar) {
      return {
        Calendar_ID: "",
        "Tipo Evento": "",
        "Tipo de Cita": "",
        ID: "",
        Nombre: "",
      };
    }
    return {
      Calendar_ID: calendar.id ? calendar.id : "",
      "Tipo de Cita": calendar.description ? calendar.description : "",
      "Tipo Evento": "",
      Nombre: calendar.summary ? calendar.summary : "",
      ID: "",
    };
  });
  try {
    const sheet = XLSX.utils.json_to_sheet(calendarsExcel);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "responses");
    XLSX.writeFile(workbook, 'data_output/tipocitas.xls')
    console.log("Libro creado ");
  } catch (err) {
    console.log("Error al querer hacer el libro " + err);
  }
};

getGoogleCalendars(TOKEN);
