using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OnlyOfficePenagihanHutang.DB;
using OnlyOfficePenagihanHutang.Models;
using OnlyOfficePenagihanHutang.Models.Home;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;

namespace OnlyOfficePenagihanHutang.Controllers
{
    public class HomeController : Controller
    {
        private string DocumentServer;
        private string DocTemplate;
        private string Folder;
        private static Random random = new Random();


        public HomeController()
        {
            DocumentServer = ConfigurationManager.AppSettings["DocumentServer"].ToString();
            DocTemplate = ConfigurationManager.AppSettings["DocTemplate"].ToString();
            Folder = "~/App_Data/";
        }

        public ActionResult Index()
        {
            var model = new HomeViewModel();

            using (var context = new AppDbContext())
            {
                model.TagihanList = context.Tagihan.ToList();
            }

            return View(model);
        }


        public ActionResult Template()
        {
            var model = new TemplateViewModel();

            var strUrlPath = $"{DocumentServer}/hosting/discovery";

            string xmlStr;
            string url = Request.Url.GetLeftPart(UriPartial.Authority);

            using (var wc = new WebClient())
            {
                xmlStr = wc.DownloadString(strUrlPath);
            }

            var xmlDoc = XDocument.Parse(xmlStr);

            var action = string.Empty;
            var ext = System.IO.Path.GetExtension(DocTemplate).ToLower().Replace(".", "");

            action = xmlDoc.Descendants("action").
                    Where(x => (String)x.Attribute("ext") == ext && (String)x.Attribute("name") == "edit").FirstOrDefault().Attribute("urlsrc").Value;

            Uri baseUri = new Uri(action);
            action = action + $"wopisrc={url}/wopi/files/{DocTemplate}";

            model.Action = action;
            return View(model);
        }


        public ActionResult Preview(string id)
        {
            var model = new TemplateViewModel();

            var strUrlPath = $"{DocumentServer}/hosting/discovery";

            string xmlStr;
            string url = Request.Url.GetLeftPart(UriPartial.Authority);

            using (var wc = new WebClient())
            {
                xmlStr = wc.DownloadString(strUrlPath);
            }

            var xmlDoc = XDocument.Parse(xmlStr);

            var action = string.Empty;
            var ext = System.IO.Path.GetExtension(DocTemplate).ToLower().Replace(".", "");

            action = xmlDoc.Descendants("action").
                    Where(x => (String)x.Attribute("ext") == ext && (String)x.Attribute("name") == "view").FirstOrDefault().Attribute("urlsrc").Value;

            Uri baseUri = new Uri(action);
            action = action + $"wopisrc={url}/wopi/files/{id}";

            model.Action = action;
            return PartialView(model);
        }



        public string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        private byte[] GenerateFile(string id)
        {
            Tagihan getTagihan;

            using (var context = new AppDbContext())
            {
                getTagihan = context.Tagihan.Where(x => x.Id == id).FirstOrDefault();
            }

            List<Tuple<string, string>> listReplacement = new List<Tuple<string, string>>();
            listReplacement.Add(new Tuple<string, string>("[Tanggal]", DateTime.Now.ToString("dd MMMM yyyy")));
            listReplacement.Add(new Tuple<string, string>("<NomerSurat>", "HUT-" + DateTime.Now.ToString("ddMMyyyy") + "-" + RandomString(5)));

            listReplacement.Add(new Tuple<string, string>("[Name]", getTagihan.Name));
            listReplacement.Add(new Tuple<string, string>("[Alamat]", getTagihan.Alamat));
            listReplacement.Add(new Tuple<string, string>("[NomerSuratHutang]", getTagihan.NomerSuratHutang));
            listReplacement.Add(new Tuple<string, string>("[Harga]", getTagihan.Harga.ToString()));
            listReplacement.Add(new Tuple<string, string>("[NomerFaktur]", getTagihan.NomerFaktur));

            var getFileTemplate = LoadFile(DocTemplate);

            byte[] getByteEditableDoc;

            using (MemoryStream stream = new MemoryStream(getFileTemplate.Bytes))
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true))
                {
                    document.ChangeDocumentType(WordprocessingDocumentType.Document);
                    var body = document.MainDocumentPart.Document.Body;

                    var runsall = body.Descendants<Run>().ToList();
                    for (int i = 0; i < runsall.Count(); i++)
                    {
                        var r = runsall[i];
                        var textsrun = r.Elements<Text>();

                        if ((listReplacement.Any(t => r.InnerText?.Trim().Contains(t.Item1) == true)))
                        {
                            var replace = listReplacement.Where(t => r.InnerText?.Trim().Contains(t.Item1) == true).FirstOrDefault();
                            foreach (var text in textsrun)
                            {
                                if (text != null)
                                {
                                    if (listReplacement.Any(t => text.InnerText?.Trim().Contains(t.Item1) == true))
                                    {
                                        var wrd = listReplacement.FirstOrDefault(it => text.InnerText?.Trim().Contains(it.Item1) == true);

                                        string[] wrdcheck = text.InnerText.Split(' ');
                                        if (wrdcheck.Count() > 0)
                                        {
                                            foreach (var key in wrdcheck)
                                            {
                                                if (listReplacement.Any(t => key?.Trim().Contains(t.Item1) == true))
                                                {
                                                    wrd = listReplacement.FirstOrDefault(it => key?.Trim().Contains(it.Item1) == true);
                                                    text.Text = text.Text.Replace(wrd.Item1, wrd.Item2);
                                                }
                                            }

                                        }
                                        else
                                        {
                                            wrd = listReplacement.FirstOrDefault(it => text.InnerText?.Trim().Contains(it.Item1) == true);
                                            text.Text = text.Text.Replace(wrd.Item1, wrd.Item2);
                                        }
                                    }

                                }
                            }

                        }

                    }

                    document.Save();
                }

                getByteEditableDoc = stream.ToArray();
            }

            return getByteEditableDoc;
        }

        public ActionResult Download(string id)
        {

            var getByteEditableDoc = GenerateFile(id);

            Tagihan getTagihan;

            using (var context = new AppDbContext())
            {
                getTagihan = context.Tagihan.Where(x => x.Id == id).FirstOrDefault();
            }

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = getTagihan.Name + DateTime.Now.ToString("ddMMyyyy")+ ".docx",
                Inline = false,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());
            Response.AppendHeader("Content-Length", getByteEditableDoc.Length.ToString());

            return File(getByteEditableDoc, "application/octet-stream");
        }

        public FileModel LoadGeneratedFile(string id)
        {
            var result = new FileModel();

            var getByteEditableDoc = GenerateFile(id);

            Tagihan getTagihan;

            using (var context = new AppDbContext())
            {
                getTagihan = context.Tagihan.Where(x => x.Id == id).FirstOrDefault();
            }

            result.Bytes = getByteEditableDoc;

            result.Mime = MimeMapping.GetMimeMapping(getTagihan.Name + DateTime.Now.ToString("ddMMyyyy") + ".docx");

            var fileInfo = new FileInfo(Folder + getTagihan.Name + DateTime.Now.ToString("ddMMyyyy") + ".docx");

            result.LastModified = fileInfo.LastWriteTime;
            result.BaseFileName = fileInfo.Name;
            result.Version = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            return result;
        }

        private FileModel LoadFile(String name)
        {

            var result = new FileModel();

            if (Folder.StartsWith("~"))
            {
                Folder = HttpContext.Server.MapPath(Folder);
            }

            result.Bytes = System.IO.File.ReadAllBytes(Folder + name);
            result.Mime = MimeMapping.GetMimeMapping(name);

            var fileInfo = new FileInfo(Folder + name);

            result.LastModified = fileInfo.LastWriteTime;
            result.BaseFileName = fileInfo.Name;
            result.Version = result.LastModified.GetValueOrDefault().ToString("yyyy-MM-dd HH:mm:ss");

            return result;
        }


        [HttpGet]
        [Route("wopi/files/{name}")]
        public ActionResult CheckFileInfo(string name)
        {
            name = name.Replace("&amp;", "&");

            FileModel file = new FileModel();

            if (name.Contains("."))
            {

                file = LoadFile(name);

                return Json(new FileInfoOnlyOffice
                {
                    BaseFileName = System.IO.Path.GetFileNameWithoutExtension(file.BaseFileName) + System.IO.Path.GetExtension(file.BaseFileName).ToLower(),
                    Version = file.Version,
                    ReadOnly = false,
                    UserCanWrite = true,
                    SupportsUpdate = true

                }, JsonRequestBehavior.AllowGet);

            }
            else
            {
                file = LoadGeneratedFile(name);

                return Json(new FileInfoOnlyOffice
                {
                    BaseFileName = System.IO.Path.GetFileNameWithoutExtension(file.BaseFileName) + System.IO.Path.GetExtension(file.BaseFileName).ToLower(),
                    Version = file.Version,
                    ReadOnly = true,
                    UserCanWrite = false,
                    SupportsUpdate = false

                }, JsonRequestBehavior.AllowGet);

            }

        }


        [HttpGet]
        [Route("wopi/files/{name}/contents")]
        public ActionResult GetFile(string name)
        {
            name = name.Replace("&amp;", "&");
            FileModel file = new FileModel();


            if (name.Contains("."))
            {
                file = LoadFile(name);
            }
            else
            {
                file = LoadGeneratedFile(name);
            }

            var cd = new System.Net.Mime.ContentDisposition
            {
                FileName = System.IO.Path.GetFileNameWithoutExtension(file.BaseFileName) + System.IO.Path.GetExtension(file.BaseFileName).ToLower(),
                Inline = false,
            };

            Response.AppendHeader("Content-Disposition", cd.ToString());
            Response.AppendHeader("Content-Length", file.Bytes.Length.ToString());

            return File(file.Bytes, "application/octet-stream");
        }

        [HttpPost]
        [Route("wopi/files/{name}/contents")]
        public ActionResult PutFile(string name)
        {

            var fileBytes = new byte[Request.InputStream.Length];
            Request.InputStream.Read(fileBytes, 0, fileBytes.Length);

            var filePath = System.IO.Path.Combine(Server.MapPath(Folder), name);
            System.IO.File.WriteAllBytes(filePath, fileBytes);

            return new HttpStatusCodeResult(HttpStatusCode.OK);

        }


        
    }
}