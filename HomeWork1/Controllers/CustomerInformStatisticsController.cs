using HomeWork1.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HomeWork1.Controllers
{
    public class CustomerInformStatisticsController : Controller
    {
        private 客戶資料Entities db = new 客戶資料Entities();
        // GET: CustomerInformStatistics
        public ActionResult Index()
        {
            var data = db.View_CustomerInformStatistics;
            return View(data);
        }
    }
}