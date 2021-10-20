using ContactManger.Models;
using System;
using System.Data;
using System.Linq;
using System.Web.Mvc;
using System.IO;
using System.Web.UI.WebControls;
using System.Web.UI;
using System.Data.Entity;
using SolrNet.Utils;

namespace ContactManger.Controllers
{
    public class ContactController : Controller
    {
        private ContactManegerDBEntities _entities = new ContactManegerDBEntities();

        // GET: /Contact/
        public ActionResult Index(string searchString)
        {
            var contacts = from c in _entities.Contacts
                           select c;
            if (!string.IsNullOrEmpty(searchString))
            {
                contacts = contacts.Where(c => c.FirstName.Contains(searchString) || c.LastName == searchString);
            }
            return View(contacts);       
        }
        //GET: /Contact/Sort/
        public ActionResult Sort(string sortOrder)
        {
            ViewBag.FirstNameSortParm = String.IsNullOrEmpty(sortOrder) ? "FirstName_desc" : "";
            ViewBag.LastNameSortParm = String.IsNullOrEmpty(sortOrder) ? "LastName_desc" : "";
            ViewBag.PhoneNoSortParm = String.IsNullOrEmpty(sortOrder) ? "PhonwNo_desc" : "";
            ViewBag.EmailIDSortParm = String.IsNullOrEmpty(sortOrder) ? "EmailID_desc" : "";
            var contacts = from c in _entities.Contacts
                           select c;
            switch (sortOrder)
            {
                case "FirstName_desc":
                    contacts = contacts.OrderByDescending(c => c.FirstName);
                    break;
                case "LastName_desc":
                    contacts = contacts.OrderByDescending(c => c.LastName);
                    break;
                case "PhoneNo_desc":
                    contacts = contacts.OrderByDescending(c => c.PhoneNo);
                    break;
                case "EmailID_desc":
                    contacts = contacts.OrderByDescending(c => c.EmailID);
                    break;
                default:
                    contacts = contacts.OrderBy(s => s.FirstName);
                    break;
            }
            return View(contacts);
        }
        // GET: /Contact/Create
        public ActionResult Create()
        {
            return View();
        }
        // POST: /Contact/Create
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Create([Bind(Exclude = "Id")] Contact contactToCreate)
        {
           
            if (!ModelState.IsValid)
                return View();
            
            try
            {
                _entities.Contacts.Add(contactToCreate);
                _entities.SaveChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        // GET: /Contact/Edit/5
        public ActionResult Edit(int id)
        {
            var contactToEdit = (from c in _entities.Contacts
                                 where c.Id == id
                                 select c).FirstOrDefault();

            return View(contactToEdit);
        }
        // POST: /Contact/Edit/5
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Edit(Contact contactToEdit)
        {
            if (!ModelState.IsValid)
                return View();
            try
            {
                var originalContact = (from c in _entities.Contacts
                                       where c.Id == contactToEdit.Id
                                       select c).FirstOrDefault();
                if (originalContact != null)
                {
                    originalContact.Id = contactToEdit.Id;
                    originalContact.FirstName = contactToEdit.FirstName;
                    originalContact.LastName = contactToEdit.LastName;
                    originalContact.PhoneNo = contactToEdit.PhoneNo;
                    originalContact.EmailID = contactToEdit.EmailID;
                    _entities.Entry(originalContact).State = System.Data.Entity.EntityState.Modified;
                    _entities.SaveChanges();
                }
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
        // GET: /Contact/Delete/5
        public ActionResult Delete(int id)
        {
            var contactToDelete = (from c in _entities.Contacts
                                   where c.Id == id
                                   select c).FirstOrDefault();
            return View(contactToDelete);
        }
        // POST: /Contact/Delete/5
        [HttpPost]
        public ActionResult Delete(Models.Contact contactToDelete)
        {
            try
            {
                var originalContact = (from c in _entities.Contacts
                                       where c.Id == contactToDelete.Id
                                       select c).FirstOrDefault();

                _entities.Contacts.Remove(originalContact);
                _entities.SaveChanges();
                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }      
        // Export To Excel Method
        public void Export()
        {
            var gv = new GridView();
            ContactManegerDBEntities entities = new ContactManegerDBEntities();
            gv.DataSource = entities.Contacts.OrderBy(x => x.Id).ToList();
            gv.DataBind();
            Response.ClearContent();
            Response.AddHeader("content-disposition",
                               string.Format("attachment;Filename = Contacts_{0}.xlsx",DateTime.Now));
            Response.ContentType = "application/excel";
            var strw = new StringWriter();
            var htmlTw = new HtmlTextWriter(strw);
            gv.RenderControl(htmlTw);
            Response.Write(strw.ToString());
            Response.End();
        }
    }
}
//COUNT 164





















//public ActionResult Search(int id)
//{
//    var contactToSearch = (from c in _entities.Contacts
//                           where c.Id == id
//                           select c).FirstOrDefault();

//    return View(contactToSearch);
//}
////POST: /Contact/Search/
//public ViewResult Search(Contact contactToSearch, string option, string search)
//{
//    var originalContact = (from c in _entities.Contacts
//                           where c.Id == contactToSearch.Id
//                           select c).FirstOrDefault();
//    //if a user choose the radio button option as Subject  
//    if (option == "FirstName")
//    {
//        return View(_entities.Contacts.Where(c => originalContact.FirstName == search || search == null).ToList());
//    }
//    else if (option == "LastName")
//    {
//        return View(_entities.Contacts.Where(c => originalContact.LastName == search || search == null).ToList());
//    }
//    else
//    {
//        return View(_entities.Contacts.Where(c => originalContact.PhoneNo == search || search == null).ToList());
//    }
//}


// Validation logic
//if (contactToCreate.FirstName.Trim().Length == 0)
//    ModelState.AddModelError("FirstName", "First name is required.");
//if (contactToCreate.LastName.Trim().Length == 0)
//    ModelState.AddModelError("LastName", "Last name is required.");
//if (contactToCreate.PhoneNo.Length > 0 && !Regex.IsMatch(contactToCreate.PhoneNo, @"((\(\d{3}\) ?)|(\d{3}-))?\d{3}-\d{4}"))
//    ModelState.AddModelError("Phone", "Invalid phone number.");
//if (contactToCreate.EmailID.Length > 0 && !Regex.IsMatch(contactToCreate.EmailID, @"^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$"))
//    ModelState.AddModelError("Email", "Invalid email address.");








////GET: /Contact/Export/
//public ActionResult Export()
//{
//   ContactManegerDBEntities entities = new ContactManegerDBEntities();
//   var contactsToExport = (from Contact in entities.Contacts.Take(10)
//                           select Contact);
//   return View(contactsToExport);
//}
////POST:/Contact/Export/
//[HttpPost]
//public ActionResult ExportToExcel(Models.Contact contactsToExport)
//{
//    ContactManegerDBEntities entities = new ContactManegerDBEntities();
//    DataTable dt = new DataTable("Grid");
//    dt.Columns.AddRange(new DataColumn[5] { new DataColumn("ContactID"), 
//                        new DataColumn("FirstName"),
//                        new DataColumn("LastName"),
//                        new DataColumn("PhoneNo"),
//                        new DataColumn("EmailID")});

//    var contacts = from Contact in entities.Contacts.Take(10)
//                         select Contact;

//    foreach (var contact in contacts)
//    {
//      dt.Rows.Add(contact.Id ,contact.FirstName ,contact.LastName ,contact.PhoneNo ,contact.EmailID);
//    }
//    using (XLWorkbook wb = new XLWorkbook())
//    {
//      wb.Worksheets.Add(dt);
//      using (MemoryStream stream = new MemoryStream())
//      {
//        wb.SaveAs(stream);
//        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Contact.xlsx");
//      }
//    }
