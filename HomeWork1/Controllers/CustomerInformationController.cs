using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HomeWork1.Models;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;

namespace HomeWork1.Controllers
{
    public class CustomerInformationController : BaseController
    {
        // GET: CustomerInformation
        public ActionResult Index(int? id)
        {
            var 客戶資料 = repo客戶資料.All();
            ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");
            return View(客戶資料.ToList());
        }
        //public ActionResult Index(int? id)
        //{
        //    ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");

        //    if (!id.HasValue)
        //    {
        //        var 客戶資料 = repo客戶資料.All();
        //        return View(客戶資料.ToList());
        //    }

        //    var 客戶資料by類別 = repo客戶資料.All().Where(p =>p.客戶分類== id);
        //    return View(客戶資料by類別.ToList());
        //}

        [HttpPost]
        public ActionResult Index(string param)
        {
            var data = repo客戶資料.All().
                Where(p => p.客戶名稱.Contains(param)
                        || p.統一編號.Contains(param)
                        || p.電話.Contains(param)
                        || p.傳真.Contains(param)
                        || p.地址.Contains(param)
                        || p.Email.Contains(param)
                        || p.客戶類別.類別.Contains(param)
                     );
            ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");
            return View(data);
        }

        public ActionResult GetCustCategoryData(string param)
        {
            var data = repo客戶資料.All();


            if (!string.IsNullOrEmpty(param))
            {
                int intParam = Convert.ToInt32(param);
                data = data.Where(p => p.客戶分類 == intParam);
            }
            
            ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");
            return View("Index", data.ToList());
        }

        public ActionResult GetExcelFile()
        {
            var 客戶資料 = repo客戶資料.All();
            MemoryStream ms = GetExportData(客戶資料);
            return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶資訊.xlsx");
        }

        private MemoryStream GetExportData(IQueryable<客戶資料> 客戶資料)
        {
            IWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("客戶資料");


            //設定表頭
            IRow rowHeader = sheet.CreateRow(0);
            rowHeader.CreateCell(0).SetCellValue("客戶名稱");
            rowHeader.CreateCell(1).SetCellValue("統一編號");
            rowHeader.CreateCell(2).SetCellValue("電話");
            rowHeader.CreateCell(3).SetCellValue("傳真");
            rowHeader.CreateCell(4).SetCellValue("地址 ");
            rowHeader.CreateCell(5).SetCellValue("Email");
            rowHeader.CreateCell(6).SetCellValue("客戶類別");

            //匯出資料
            int count = 1;
            foreach (var item in 客戶資料)
            {
                IRow row = sheet.CreateRow(count);
                row.CreateCell(0).SetCellValue(item.客戶名稱);
                row.CreateCell(1).SetCellValue(item.統一編號);
                row.CreateCell(2).SetCellValue(item.電話);
                row.CreateCell(3).SetCellValue(item.傳真);
                row.CreateCell(4).SetCellValue(item.地址);
                row.CreateCell(5).SetCellValue(item.Email);
                row.CreateCell(6).SetCellValue(item.客戶類別.類別);
                count++;
            }

            workbook.Write(ms);
            workbook = null;
            return ms;
        }

        public ActionResult CustContactsPartialView()
        {
            var 客戶聯絡人 = repo客戶聯絡人.All().Include(客 => 客.客戶資料);

            return View(客戶聯絡人.ToList());
        }

        // GET: CustomerInformation/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        [HttpPost]
        public ActionResult Details(IList<BatchUpdateCustContacts> CustContactsList)
        {

            if (ModelState.IsValid)
            {
                foreach (var item in CustContactsList)
                {
                    var 客戶聯絡人 = repo客戶聯絡人.Find(item.Id);
                    客戶聯絡人.職稱 = item.職稱;
                    客戶聯絡人.手機 = item.手機;
                    客戶聯絡人.電話 = item.電話;
                }
                repo客戶聯絡人.UnitOfWork.Commit();

                return RedirectToAction("Index");
            }
            客戶資料 客戶資料 = repo客戶聯絡人.Find(CustContactsList.FirstOrDefault().Id).客戶資料;
            return View(客戶資料);
        }

        // GET: CustomerInformation/Create
        public ActionResult Create()
        {
            ViewBag.客戶分類 = new SelectList(repo客戶類別.All(), "Id", "類別");
            return View();
        }

        // POST: CustomerInformation/Create
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,客戶分類")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                repo客戶資料.Add(客戶資料);
                repo客戶資料.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }
            ViewBag.客戶分類 = new SelectList(repo客戶類別.All(), "Id", "類別");
            return View(客戶資料);
        }

        // GET: CustomerInformation/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            ViewBag.客戶分類 = new SelectList(repo客戶類別.All(), "Id", "類別", 客戶資料.客戶分類);
            return View(客戶資料);
        }

        // POST: CustomerInformation/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,客戶名稱,統一編號,電話,傳真,地址,Email,客戶分類")] 客戶資料 客戶資料)
        {
            if (ModelState.IsValid)
            {
                var db客戶資料 = repo客戶資料.UnitOfWork.Context;
                db客戶資料.Entry(客戶資料).State = EntityState.Modified;
                db客戶資料.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.客戶分類 = new SelectList(repo客戶類別.All(), "Id", "類別", 客戶資料.客戶分類);
            return View(客戶資料);
        }

        // GET: CustomerInformation/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            客戶資料 客戶資料 = repo客戶資料.Find(id.Value);
            if (客戶資料 == null)
            {
                return HttpNotFound();
            }
            return View(客戶資料);
        }

        // POST: CustomerInformation/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            客戶資料 客戶資料 = repo客戶資料.Find(id);
            客戶資料.是否已刪除 = true;

            repo客戶資料.Delete(客戶資料);
            repo客戶資料.UnitOfWork.Commit();

            return RedirectToAction("Index");
        }

        
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                repo客戶資料.UnitOfWork.Context.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
