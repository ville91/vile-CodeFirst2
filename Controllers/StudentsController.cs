using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Vile_CodeFirst_2.DBContext;
using Vile_CodeFirst_2.Models;

namespace Vile_CodeFirst_2.Controllers
{
    public class StudentsController : Controller
    {
        private StudentManagementDBContext db = new StudentManagementDBContext();

        // GET: Students
        public ActionResult Index()
        {
            return View(db.Students.ToList());
        }

        // GET: Students/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Student student = db.Students.Find(id);
            if (student == null)
            {
                return HttpNotFound();
            }
            return View(student);
        }

        // GET: Students/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Students/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Name,birthdate")] Student student)
        {
            if (ModelState.IsValid)
            {
                db.Students.Add(student);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(student);
        }

        // GET: Students/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Student student = db.Students.Find(id);
            if (student == null)
            {
                return HttpNotFound();
            }
            return View(student);
        }

        // POST: Students/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Name,birthdate")] Student student)
        {
            if (ModelState.IsValid)
            {
                db.Entry(student).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(student);
        }

        // GET: Students/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Student student = db.Students.Find(id);
            if (student == null)
            {
                return HttpNotFound();
            }
            return View(student);
        }

        // POST: Students/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Student student = db.Students.Find(id);
            db.Students.Remove(student);
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        [HttpGet]
        public ActionResult Export()
        {
            // Create the workbook and worksheet
            var wb = new XLWorkbook();
            var studentSheet = wb.AddWorksheet("Students");
            var currentRow = 1;

            // Add first row header
            studentSheet.Cell(currentRow, 1).Value = "Id";
            studentSheet.Cell(currentRow, 2).Value = "Name";
            studentSheet.Cell(currentRow, 3).Value = "Birthdate";

            // Get all student from database
            List<Student> allStudent = db.Students.ToList();
            foreach (var student in allStudent)
            {
                currentRow++;
                studentSheet.Cell(currentRow, 1).Value = student.ID;
                studentSheet.Cell(currentRow, 2).Value = student.Name;
                studentSheet.Cell(currentRow, 3).Value = student.birthdate;
            }

            // Convert workbook to byte array then response to client
            using (var memoryStream = new MemoryStream())
            {
                wb.SaveAs(memoryStream);
                var byteArrayContent = memoryStream.ToArray();
                var responseResult = File(byteArrayContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "all-students.xlsx");
                return responseResult;
            }
        }
        [HttpGet]
        public ActionResult Import()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase file)
        {
            // Validate file name
            if (file.FileName.EndsWith("xlsx"))
            {
                var wb = new XLWorkbook(file.InputStream);
                if (wb.TryGetWorksheet("Students", out var wSheet))
                {
                    var allRow = wSheet.Rows();
                    var rowIndex = 0;
                    foreach (var row in allRow)
                    {
                        if (rowIndex > 0)
                        {
                            var allCellOfRow = row.Cells().ToArray();
                            var name = allCellOfRow[1].Value.ToString();
                            var birthdate = int.Parse(allCellOfRow[2].Value.ToString());
                            var newStudent = new Student
                            {
                                Name = name,
                                birthdate = birthdate
                            };
                            db.Students.Add(newStudent);

                        }
                        rowIndex++;
                    }
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
            }
            return RedirectToAction("NotFound");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
        public HttpStatusCodeResult NotFound()
        {
            return new HttpStatusCodeResult(System.Net.HttpStatusCode.NotFound);
        }

    }
}
