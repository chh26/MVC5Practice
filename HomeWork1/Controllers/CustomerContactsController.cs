using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HomeWork1.Models;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace HomeWork1.Controllers
{
    public class CustomerContactsController : BaseController
    {
        // GET: CustomerContacts
        public ActionResult Index()
        {
            var 客戶聯絡人 = repo客戶聯絡人.All().Include(客 => 客.客戶資料);

            return View(客戶聯絡人.ToList());
        }

        [HttpPost]
        public ActionResult Index(string param)
        {
            var data = repo客戶聯絡人.All().
                Where(p => p.職稱.Contains(param)
                        || p.姓名.ToString().Contains(param)
                        || p.Email.ToString().Contains(param)
                        || p.手機.Contains(param)
                        || p.電話.Contains(param)
                     ).AsQueryable();
            return View(data);
        }

        public ActionResult PartialViewTest()
        {
            //不載入Layout
            return PartialView("Index");
        }

        public ActionResult GetExcelFile()
        {
            var 客戶聯絡人 = repo客戶聯絡人.All();
            MemoryStream ms = GetExportData(客戶聯絡人);
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶聯絡人資訊.xlsx");
        }

        private MemoryStream GetExportData(IQueryable<客戶聯絡人> 客戶聯絡人)
        {
            IWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("客戶聯絡人資訊");


            //設定表頭
            IRow rowHeader = sheet.CreateRow(0);
            rowHeader.CreateCell(0).SetCellValue("職稱");
            rowHeader.CreateCell(1).SetCellValue("姓名");
            rowHeader.CreateCell(2).SetCellValue("Email");
            rowHeader.CreateCell(3).SetCellValue("手機");
            rowHeader.CreateCell(4).SetCellValue("電話 ");
            rowHeader.CreateCell(5).SetCellValue("客戶名稱");

            //匯出資料
            int count = 1;
            foreach (var item in 客戶聯絡人)
            {
                IRow row = sheet.CreateRow(count);
                row.CreateCell(0).SetCellValue(item.職稱);
                row.CreateCell(1).SetCellValue(item.姓名);
                row.CreateCell(2).SetCellValue(item.Email);
                row.CreateCell(3).SetCellValue(item.手機);
                row.CreateCell(4).SetCellValue(item.電話);
                row.CreateCell(5).SetCellValue(item.客戶資料.客戶名稱);
                count++;
            }

            workbook.Write(ms);
            workbook = null;
            return ms;
        }

        // GET: CustomerContacts/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // GET: CustomerContacts/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: CustomerContacts/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                repo客戶聯絡人.Add(客戶聯絡人);
                repo客戶聯絡人.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: CustomerContacts/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // POST: CustomerContacts/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,職稱,姓名,Email,手機,電話")] 客戶聯絡人 客戶聯絡人)
        {
            if (ModelState.IsValid)
            {
                var db客戶聯絡人 = repo客戶聯絡人.UnitOfWork.Context;
                db客戶聯絡人.Entry(客戶聯絡人).State = EntityState.Modified;
                db客戶聯絡人.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶聯絡人.客戶Id);
            return View(客戶聯絡人);
        }

        // GET: CustomerContacts/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id.Value);
            if (客戶聯絡人 == null)
            {
                return HttpNotFound();
            }
            return View(客戶聯絡人);
        }

        // POST: CustomerContacts/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶聯絡人 客戶聯絡人 = repo客戶聯絡人.Find(id);
            repo客戶聯絡人.Delete(客戶聯絡人);
            repo客戶聯絡人.UnitOfWork.Commit();

            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo客戶聯絡人.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
