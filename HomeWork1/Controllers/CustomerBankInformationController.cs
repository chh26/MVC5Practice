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
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace HomeWork1.Controllers
{
    public class CustomerBankInformationController : BaseController
    {
        // GET: CustomerBankInformation
        public ActionResult Index()
        {
            var 客戶銀行資訊 = repo客戶銀行資訊.All();

            return View(客戶銀行資訊.ToList());
        }

        [HttpPost]
        public ActionResult Index(string param)
        {
            var data = repo客戶銀行資訊.All().
                Where(p => p.銀行名稱.Contains(param)
                        || p.銀行代碼.ToString().Contains(param)
                        || p.分行代碼.ToString().Contains(param)
                        || p.帳戶名稱.Contains(param)
                        || p.帳戶號碼.Contains(param)
                        || p.客戶資料.客戶名稱.Contains(param)
                     ).AsQueryable();
            TempData["param"] = param;
            return View(data);
        }

        public ActionResult GetExcelFile(string hidparam)
        {
            var 客戶銀行資訊 = repo客戶銀行資訊.All().
                Where(p => p.銀行名稱.Contains(hidparam)
                        || p.銀行代碼.ToString().Contains(hidparam)
                        || p.分行代碼.ToString().Contains(hidparam)
                        || p.帳戶名稱.Contains(hidparam)
                        || p.帳戶號碼.Contains(hidparam)
                        || p.客戶資料.客戶名稱.Contains(hidparam)
                     ).AsQueryable();
            MemoryStream ms = GetExportData(客戶銀行資訊);
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶銀行資訊.xlsx");
        }

        private MemoryStream GetExportData(IQueryable<客戶銀行資訊> 客戶銀行資訊)
        {
            IWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("客戶銀行資訊");


            //設定表頭
            IRow rowHeader = sheet.CreateRow(0);
            rowHeader.CreateCell(0).SetCellValue("銀行名稱");
            rowHeader.CreateCell(1).SetCellValue("銀行代碼");
            rowHeader.CreateCell(2).SetCellValue("分行代碼");
            rowHeader.CreateCell(3).SetCellValue("帳戶名稱");
            rowHeader.CreateCell(4).SetCellValue("帳戶號碼 ");
            rowHeader.CreateCell(5).SetCellValue("客戶名稱");

            //匯出資料
            int count = 1;
            foreach (var item in 客戶銀行資訊)
            {
                IRow row = sheet.CreateRow(count);
                row.CreateCell(0).SetCellValue(item.銀行名稱);
                row.CreateCell(1).SetCellValue(item.銀行代碼);
                row.CreateCell(2).SetCellValue(item.分行代碼.ToString());
                row.CreateCell(3).SetCellValue(item.帳戶名稱);
                row.CreateCell(4).SetCellValue(item.帳戶號碼);
                row.CreateCell(5).SetCellValue(item.客戶資料.客戶名稱);
                count++;
            }

            workbook.Write(ms);
            workbook = null;
            return ms;
        }

        // GET: CustomerBankInformation/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // GET: CustomerBankInformation/Create
        public ActionResult Create()
        {
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱");
            return View();
        }

        // POST: CustomerBankInformation/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                repo客戶銀行資訊.Add(客戶銀行資訊);
                repo客戶銀行資訊.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: CustomerBankInformation/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // POST: CustomerBankInformation/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶Id,銀行名稱,銀行代碼,分行代碼,帳戶名稱,帳戶號碼")] 客戶銀行資訊 客戶銀行資訊)
        {
            if (ModelState.IsValid)
            {
                var db客戶銀行資訊 = repo客戶銀行資訊.UnitOfWork.Context;
                db客戶銀行資訊.Entry(客戶銀行資訊).State = EntityState.Modified;
                db客戶銀行資訊.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.客戶Id = new SelectList(repo客戶資料.All(), "Id", "客戶名稱", 客戶銀行資訊.客戶Id);
            return View(客戶銀行資訊);
        }

        // GET: CustomerBankInformation/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id.Value);
            if (客戶銀行資訊 == null)
            {
                return HttpNotFound();
            }
            return View(客戶銀行資訊);
        }

        // POST: CustomerBankInformation/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶銀行資訊 客戶銀行資訊 = repo客戶銀行資訊.Find(id);
            repo客戶銀行資訊.Delete(客戶銀行資訊);
            repo客戶銀行資訊.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo客戶銀行資訊.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
