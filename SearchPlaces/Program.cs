// API key de Google Places


using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Web;
using static System.Net.Mime.MediaTypeNames;
using ClosedXML.Excel;


string apiKey = "AIzaSyALmmDJaBTq1yQKxPVwZE359skx6ricbFA";

// Lugar a buscar
string placeName = "Pizzeria";

// Zona a buscar
string location = "-34.660216,-58.548723"; // latitud y longitud

// Radio de búsqueda (en metros)
int radius = 5000;

// URL de la API de Google Places
string apiUrl = string.Format("https://maps.googleapis.com/maps/api/place/nearbysearch/json?location={0}&radius={1}&name={2}&key={3}", HttpUtility.UrlEncode(location), radius, HttpUtility.UrlEncode(placeName), apiKey);

// Descarga los datos de la API
WebClient client = new WebClient();
string jsonResult = client.DownloadString(apiUrl);

// Convierte los datos a objetos JSON
JObject result = JObject.Parse(jsonResult);

// Iterar sobre los resultados y obtener nombres y teléfonos
JArray results = (JArray)result["results"];

var workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("Datos");

// Escribir encabezados
worksheet.Cell(1, 1).Value = "Nombre";
worksheet.Cell(1, 2).Value = "Teléfono";
worksheet.Cell(1, 3).Value = "Dirección";
worksheet.Cell(1, 4).Value = "Sitio web";
worksheet.Cell(1, 5).Value = "URL de foto";

// Escribir datos
int row = 2;

    foreach (var place in results)
{
    string placeId = place["place_id"].ToString();

    // Obtener detalles del lugar
    apiUrl = string.Format("https://maps.googleapis.com/maps/api/place/details/json?place_id={0}&key={1}&fields=name,formatted_phone_number,formatted_address,website,opening_hours,photos", placeId, apiKey);
    jsonResult = client.DownloadString(apiUrl);
    JObject placeDetails = JObject.Parse(jsonResult);

    // Obtener el nombre, número de teléfono, dirección y la imagen del lugar
    string name = placeDetails["result"]["name"].ToString();
    string phoneNumber = placeDetails["result"]["formatted_phone_number"]?.ToString() ?? "No disponible";
    string address = placeDetails["result"]["formatted_address"]?.ToString() ?? "No disponible";
    string website = placeDetails["result"]["website"]?.ToString() ?? "No disponible";
    JArray photos = placeDetails["result"]["photos"] as JArray;
    string photoUrl = photos != null && photos.Count > 0 ? string.Format("https://maps.googleapis.com/maps/api/place/photo?maxwidth=400&photoreference={0}&key={1}", photos[0]["photo_reference"], apiKey) : "No disponible";

        worksheet.Cell(row, 1).Value = name;
        worksheet.Cell(row, 2).Value = phoneNumber;
        worksheet.Cell(row, 3).Value = address;
        worksheet.Cell(row, 4).Value = website;
        worksheet.Cell(row, 5).Value = photoUrl;
        row++;

        Console.WriteLine(name);
    Console.WriteLine(phoneNumber);
    Console.WriteLine(address);
    Console.WriteLine(website);
    Console.WriteLine(photoUrl);
}
//var file = new FileInfo("datos.xlsx");
var file = new FileInfo(@"C:\Usuarios\Alexis\Desktop\datos.xlsx");
workbook.SaveAs(file.FullName);
Console.ReadLine();
