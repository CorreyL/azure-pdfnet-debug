using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web.Http;
using pdftron;
using System.Diagnostics;
using pdftron.PDF;

namespace PDF2Word.Controllers
{
	[ApiController]
	[Microsoft.AspNetCore.Mvc.Route("[controller]")]
	public class WeatherForecastController : ControllerBase
	{
		private static readonly string[] Summaries = new[]
		{
			"Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
		};

		private readonly ILogger<WeatherForecastController> _logger;

		public WeatherForecastController(ILogger<WeatherForecastController> logger)
		{
			_logger = logger;
		}

		private string ConvertPDF2Word() {
			pdftron.PDF.Convert.WordOutputOptions wordOutputOptions = new pdftron.PDF.Convert.WordOutputOptions();

			// Optionally convert only the first page
			wordOutputOptions.SetPages(1, 1);

			// Requires the Structured Output module
			pdftron.PDF.Convert.ToWord("./SYH_Letter_red.docx.pdf", "./test.docx", wordOutputOptions);
			byte[] docx = System.IO.File.ReadAllBytes("./test.docx");
			HttpResponseMessage responseMsg = new HttpResponseMessage(HttpStatusCode.OK);
			responseMsg.Content = new ByteArrayContent(docx);
			responseMsg.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
			responseMsg.Content.Headers.ContentDisposition.FileName = "test.docx";
			responseMsg.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
			return System.Convert.ToBase64String(docx);
		}

		private string Office2PDF(bool loggingEnabled) {
			using var pdfdoc = new pdftron.PDF.PDFDoc();

			FileStream docxFile = System.IO.File.OpenRead("./test.docx");
			byte[] byteArray = new byte[docxFile.Length];
			docxFile.Read(byteArray, 0, byteArray.Length);
			pdftron.Filters.MemoryFilter memoryFilter = new pdftron.Filters.MemoryFilter(byteArray.Length, false);
			pdftron.Filters.FilterWriter filterWriter = new pdftron.Filters.FilterWriter(memoryFilter);

			filterWriter.WriteBuffer(byteArray);
			filterWriter.Flush();
			memoryFilter.SetAsInputFilter();

			Stopwatch sw = Stopwatch.StartNew();

			pdftron.PDF.Convert.OfficeToPDF(pdfdoc, memoryFilter, null);  //the operation in question which has such a huge difference locally vs running on Azure
			Console.WriteLine($"Office to pdf duration: {sw.ElapsedMilliseconds} ms");

			sw.Stop();
			return $"Office to pdf duration: {sw.ElapsedMilliseconds} ms and logging is: {loggingEnabled}";
		}

		private string DocumentTemplating(bool loggingEnabled) {
			Stopwatch sw = Stopwatch.StartNew();

			TemplateDocument template_doc = pdftron.PDF.Convert.CreateOfficeTemplate("letter_template.docx", null);
			String json = "{\"company_name\": \"PDFTron\", \"Date\": \"Test\", \"applicant_first_name\": \"Dale\", \"applicant_surname\": \"Cooper\"}";
			// Fill the template with data from a JSON string, producing a PDF document.
			PDFDoc pdf_doc = template_doc.FillTemplateJson(json);

			// Save the PDF to memory.
			pdf_doc.Save("./officetemplate.pdf", pdftron.SDF.SDFDoc.SaveOptions.e_linearized);

			sw.Stop();

			return $"Office Templating duration: {sw.ElapsedMilliseconds} ms and logging is: {loggingEnabled}";
		}

		[System.Web.Http.HttpGet]
		public String Get()
		{
			Console.WriteLine($"{Directory.GetCurrentDirectory()}");
			pdftron.PDFNet.Initialize("");
			pdftron.PDFNet.AddResourceSearchPath("./");
			bool loggingEnabled = false;
			if (PDFNetInternalTools.IsLogSystemAvailable())
			{
				loggingEnabled = true;
				PDFNetInternalTools.SetDefaultLogThreshold(PDFNetInternalToolsLogLevel.e_pdf_net_internal_tools_trace);

				PDFNetInternalTools.SetLogLocation("./", "SendToPDFTronSupport.log.txt"); // Set to your folder

				PDFNetInternalTools.DisableLogBackend(PDFNetInternalToolsLogBackend.e_pdf_net_internal_tools_callback);
				PDFNetInternalTools.DisableLogBackend(PDFNetInternalToolsLogBackend.e_pdf_net_internal_tools_debugger);
				PDFNetInternalTools.DisableLogBackend(PDFNetInternalToolsLogBackend.e_pdf_net_internal_tools_console);
				PDFNetInternalTools.EnableLogBackend(PDFNetInternalToolsLogBackend.e_pdf_net_internal_tools_disk);
			}
			// return ConvertPDF2Word();
			// return Office2PDF(loggingEnabled);
			return DocumentTemplating(loggingEnabled);
		}
	}
}
