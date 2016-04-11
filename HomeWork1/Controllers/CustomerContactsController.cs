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
using System.Linq.Expressions;
using PagedList;

namespace HomeWork1.Controllers
{
    public class CustomerContactsController : BaseController
    {
        private int pageSize = 1;
        // GET: CustomerContacts
        public ActionResult Index(int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var 客戶聯絡人 = repo客戶聯絡人.All();

            ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct());
            return View(客戶聯絡人.ToList().ToPagedList(currentPage, pageSize));
        }

        [HttpPost]
        public ActionResult Index(string param, string hidjobtitle, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var data = repo客戶聯絡人.Search(param, hidjobtitle);

            if (!string.IsNullOrEmpty(hidjobtitle))
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct(), hidjobtitle);
            else
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct());

            TempData["param"] = param;
            TempData["jobtitle"] = hidjobtitle;

            return View(data.ToList().ToPagedList(currentPage, pageSize));
        }

        [HandleError(ExceptionType = typeof(ArgumentException), View = "SearchError")]
        public ActionResult GetJobList(string hidparam, string jobtitle, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            if (hidparam == null && jobtitle == null)
            {
                throw new ArgumentException("參數錯誤");
            }
            var data = repo客戶聯絡人.Search(hidparam, jobtitle);

            if (!string.IsNullOrEmpty(jobtitle))
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct(), jobtitle);
            else
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct());

            TempData["param"] = hidparam;
            TempData["jobtitle"] = jobtitle;
            return View("Index", data.ToList().ToPagedList(currentPage, pageSize));
        }

        public ActionResult PartialViewTest()
        {
            //不載入Layout
            return PartialView("Index");
        }

        [HttpPost]
        public ActionResult GetExcelFile(string hidparam, string hidjobtitle)
        {
            var 客戶聯絡人 = repo客戶聯絡人.Search(hidparam, hidjobtitle);

            TempData["param"] = hidparam;
            TempData["jobtitle"] = hidjobtitle;

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

        public ActionResult Sort(string columnName, string sort, string sortColumn, string keyword, string jobtitle, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var data = repo客戶聯絡人.Search(keyword, jobtitle);

            if (columnName == "客戶名稱")
            {
                if (columnName == sortColumn)
                {
                    if (sort == "Desc")
                    {
                        data = data.OrderByDescending(p => p.客戶資料.客戶名稱);
                        ViewBag.sortTag = "▼";
                        TempData["nextSort"] = "Asc";
                        TempData["pageSort"] = "Desc";
                    }
                    else
                    {
                        data = data.OrderBy(p => p.客戶資料.客戶名稱);
                        ViewBag.sortTag = "▲";
                        TempData["nextSort"] = "Desc";
                        TempData["pageSort"] = "Asc";
                    }
                }
                else
                {
                    data = data.OrderBy(p => p.客戶資料.客戶名稱);
                    ViewBag.sortTag = "▲";
                    TempData["nextSort"] = "Desc";
                    TempData["pageSort"] = "Asc";
                }
            }
            else
            {
                var param = Expression.Parameter(typeof(客戶聯絡人), "custContacts");
                var orderExpression = Expression.Lambda<Func<客戶聯絡人, object>>(Expression.Property(param, columnName), param);
                if (columnName == sortColumn)
                {
                    if (sort == "Desc")
                    {
                        data = data.OrderByDescending(orderExpression);
                        ViewBag.sortTag = "▼";
                        TempData["nextSort"] = "Asc";
                        TempData["pageSort"] = "Desc";
                    }
                    else
                    {
                        data = data.OrderBy(orderExpression);
                        ViewBag.sortTag = "▲";
                        TempData["nextSort"] = "Desc";
                        TempData["pageSort"] = "Asc";
                    }
                }
                else
                {
                    data = data.OrderBy(orderExpression);
                    ViewBag.sortTag = "▲";
                    TempData["nextSort"] = "Desc";
                    TempData["pageSort"] = "Asc";
                }
            }


            if (!string.IsNullOrEmpty(jobtitle))
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct(), jobtitle);
            else
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct());

            TempData["sortColumn"] = columnName;
            TempData["param"] = keyword;
            TempData["jobtitle"] = jobtitle;
            TempData["currentPage"] = currentPage;

            return View("Index", data.ToList().ToPagedList(currentPage, pageSize));
        }

        public ActionResult PageList(string sort, string sortColumn, string keyword, string jobtitle, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var data = repo客戶聯絡人.Search(keyword, jobtitle);

            if (!string.IsNullOrEmpty(sortColumn))
            {
                if (sortColumn == "客戶分類")
                {
                    if (sort == "Desc")
                    {
                        data = data.OrderByDescending(p => p.客戶資料.客戶名稱);
                        ViewBag.sortTag = "▼";
                        TempData["nextSort"] = "Asc";
                        TempData["pageSort"] = "Desc";
                    }
                    else
                    {
                        data = data.OrderBy(p => p.客戶資料.客戶名稱);
                        ViewBag.sortTag = "▲";
                        TempData["nextSort"] = "Desc";
                        TempData["pageSort"] = "Asc";
                    }
                }
                else
                {
                    var param = Expression.Parameter(typeof(客戶聯絡人), "custContacts");
                    var orderExpression = Expression.Lambda<Func<客戶聯絡人, object>>(Expression.Property(param, sortColumn), param);
                    if (sort == "Desc")
                    {
                        data = data.OrderByDescending(orderExpression);
                        ViewBag.sortTag = "▼";
                        TempData["nextSort"] = "Asc";
                        TempData["pageSort"] = "Desc";
                    }
                    else
                    {
                        data = data.OrderBy(orderExpression);
                        ViewBag.sortTag = "▲";
                        TempData["nextSort"] = "Desc";
                        TempData["pageSort"] = "Asc";
                    }
                }
            }

            if (!string.IsNullOrEmpty(jobtitle))
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct(), jobtitle);
            else
                ViewBag.jobTitleList = new SelectList(repo客戶聯絡人.All().Select(p => p.職稱).Distinct());

            TempData["sortColumn"] = sortColumn;
            TempData["param"] = keyword;
            TempData["jobtitle"] = jobtitle;
            TempData["currentPage"] = currentPage;

            return View("Index", data.ToList().ToPagedList(currentPage, pageSize));
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
