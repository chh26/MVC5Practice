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
using System.Web.Security;
using System.Linq.Expressions;
using PagedList;

namespace HomeWork1.Controllers
{
    public class CustomerInformationController : BaseController
    {
        private int pageSize = 1;
        // GET: CustomerInformation
        public ActionResult Index(int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var 客戶資料 = repo客戶資料.All();
            ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");
            return View(客戶資料.ToList().ToPagedList(currentPage, pageSize));
        }
        
        [HttpPost]
        public ActionResult Index(string param,string hidcustCategory, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var data = repo客戶資料.Search(param, hidcustCategory);

            if (!string.IsNullOrEmpty(hidcustCategory))
            {
                int intCategory = Convert.ToInt32(hidcustCategory);
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別", intCategory);
            }
            else
            {
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");
            }

            TempData["param"] = param;
            TempData["custCategory"] = hidcustCategory;
            return View(data.ToList().ToPagedList(currentPage, pageSize));
        }

        [HandleError(ExceptionType = typeof(ArgumentException), View = "SearchError")]
        public ActionResult GetCustCategoryData(string hidparam, string custCategory, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            if (hidparam == null && custCategory == null)
            {
                throw new ArgumentException("參數錯誤");
            }

            var data = repo客戶資料.Search(hidparam, custCategory);

            if (!string.IsNullOrEmpty(custCategory))
            {
                int intParam = Convert.ToInt32(custCategory);
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別", intParam);
            }
            else
            {
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");
            }

            TempData["param"] = hidparam;
            TempData["custCategory"] = custCategory;
            return View("Index", data.ToList().ToPagedList(currentPage, pageSize));
        }

        [HttpPost]
        public ActionResult GetExcelFile(string hidparam, string hidcustCategory)
        {
            var 客戶資料 = repo客戶資料.Search(hidparam, hidcustCategory);

            if (!string.IsNullOrEmpty(hidcustCategory))
            {
                int intCategory = Convert.ToInt32(hidcustCategory);
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別", intCategory);
            }
            else
            {
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");
            }

            TempData["param"] = hidparam;
            TempData["custCategory"] = hidcustCategory;
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

        public ActionResult Sort(string columnName, string sort, string sortColumn, string keyword, string CategoryId, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var data = repo客戶資料.Search(keyword, CategoryId);

            if (columnName == "客戶分類")
            {
                if (columnName == sortColumn)
                {
                    if (sort == "Desc")
                    {
                        data = data.OrderByDescending(p => p.客戶類別.類別);
                        ViewBag.sortTag = "▼";
                        TempData["nextSort"] = "Asc";
                        TempData["pageSort"] = "Desc";
                    }
                    else
                    {
                        data = data.OrderBy(p => p.客戶類別.類別);
                        ViewBag.sortTag = "▲";
                        TempData["nextSort"] = "Desc";
                        TempData["pageSort"] = "Asc";
                    }
                }
                else
                {
                    data = data.OrderBy(p => p.客戶類別.類別);
                    ViewBag.sortTag = "▲";
                    TempData["nextSort"] = "Desc";
                    TempData["pageSort"] = "Asc";
                }
            }
            else
            {
                var param = Expression.Parameter(typeof(客戶資料), "customerinfo");
                var orderExpression = Expression.Lambda<Func<客戶資料, object>>(Expression.Property(param, columnName), param);
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

           
            if (!string.IsNullOrEmpty(CategoryId))
            {
                int intCategory = Convert.ToInt32(CategoryId);
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別", intCategory);
            }
            else
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");

            TempData["sortColumn"] = columnName;
            TempData["param"] = keyword;
            TempData["custCategory"] = CategoryId;
            TempData["currentPage"] = currentPage;

            return View("Index", data.ToList().ToPagedList(currentPage, pageSize));
        }

        public ActionResult PageList(string sort, string sortColumn, string keyword, string CategoryId, int page = 1)
        {
            int currentPage = page > 1 ? page : 1;
            var data = repo客戶資料.Search(keyword, CategoryId);

            if (!string.IsNullOrEmpty(sortColumn))
            {
                if (sortColumn == "客戶分類")
                {
                    if (sort == "Desc")
                    {
                        data = data.OrderByDescending(p => p.客戶類別.類別);
                        ViewBag.sortTag = "▼";
                        TempData["nextSort"] = "Asc";
                        TempData["pageSort"] = "Desc";
                    }
                    else
                    {
                        data = data.OrderBy(p => p.客戶類別.類別);
                        ViewBag.sortTag = "▲";
                        TempData["nextSort"] = "Desc";
                        TempData["pageSort"] = "Asc";
                    }
                }
                else
                {
                    var param = Expression.Parameter(typeof(客戶資料), "customerinfo");
                    var orderExpression = Expression.Lambda<Func<客戶資料, object>>(Expression.Property(param, sortColumn), param);
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

            if (!string.IsNullOrEmpty(CategoryId))
            {
                int intCategory = Convert.ToInt32(CategoryId);
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別", intCategory);
            }
            else
                ViewBag.custCategory = new SelectList(repo客戶類別.All(), "Id", "類別");

            TempData["sortColumn"] = sortColumn;
            TempData["param"] = keyword;
            TempData["custCategory"] = CategoryId;
            TempData["currentPage"] = currentPage;

            return View("Index", data.ToList().ToPagedList(currentPage, pageSize));
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
            //TempData["pwd"] = 客戶資料.密碼;
            ViewBag.客戶分類 = new SelectList(repo客戶類別.All(), "Id", "類別", 客戶資料.客戶分類);
            return View(客戶資料);
        }

        // POST: CustomerInformation/Edit/5
        // 若要免於過量張貼攻擊，請啟用想要繫結的特定屬性，如需
        // 詳細資訊，請參閱 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, FormCollection form)
        {
            var 客戶資料 = repo客戶資料.Find(id);
            if (TryUpdateModel(客戶資料, new string[] {
                "密碼","電話","傳真","地址","Email"}))
            {
                var AA = form["PWD"];
                if (!string.IsNullOrEmpty(AA))
                    客戶資料.密碼 = FormsAuthentication.HashPasswordForStoringInConfigFile(AA, "SHA1");
                else
                    客戶資料.密碼 = form["hidpwd"];

                repo客戶資料.UnitOfWork.Commit();
                TempData["msg"] = "更新成功。";
                //return RedirectToAction("Index");
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
