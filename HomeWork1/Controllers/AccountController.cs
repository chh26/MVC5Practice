using HomeWork1.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace HomeWork1.Controllers
{
    [Authorize]
    public class AccountController : BaseController
    {
        string userData = "";

        [AllowAnonymous]
        public ActionResult Login()
        {
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        public ActionResult Login(客戶資料 data)
        {

            // 登入時清空所有 Session 資料
            Session.RemoveAll();

            // 登入的密碼（以 SHA1 加密）
            string strPassword = string.Empty;
            if (!string.IsNullOrEmpty(data.密碼))
                strPassword= FormsAuthentication.HashPasswordForStoringInConfigFile(data.密碼, "SHA1");

            if (ValidateLogin(data.帳號, strPassword))
            {
                // 將管理者登入的 Cookie 設定成 Session Cookie
                bool isPersistent = false;

                FormsAuthenticationTicket ticket = new FormsAuthenticationTicket(1,
                  data.帳號,
                  DateTime.Now,
                  DateTime.Now.AddMinutes(30),
                  isPersistent,
                  userData,
                  FormsAuthentication.FormsCookiePath);

                string encTicket = FormsAuthentication.Encrypt(ticket);

                // Create the cookie.
                Response.Cookies.Add(new HttpCookie(FormsAuthentication.FormsCookieName, encTicket));

                ViewBag.authority = userData;

                var repo = repo客戶資料.Where(p => p.帳號 == data.帳號).FirstOrDefault();
                ViewBag.客戶分類 = new SelectList(repo客戶類別.All(), "Id", "類別", repo.客戶分類);
                ViewBag.Id = repo.Id;
                return RedirectToAction("Edit", "CustomerInformation", new { id = repo.Id });
            }
           
            return View();
        }

        private bool ValidateLogin(string account, string pwd)
        {
            var repo = repo客戶資料.Where(p => p.帳號 == account).FirstOrDefault();
            if (repo != null)
            {
                if (string.IsNullOrEmpty(repo.密碼) && string.IsNullOrEmpty(pwd))
                {
                    if (account == "taylor")
                        userData = "admin";
                    else
                        userData = "customer";

                    return true;
                }
                else if (repo.密碼 == pwd)
                {
                    if (account == "taylor")
                        userData = "admin";
                    else
                        userData = "customer";

                    return true;
                }
                else
                    return false;
            }
            
            return false;
        }

        [AllowAnonymous]
        public ActionResult Register()
        {
            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        public ActionResult Register(RegisterViewModel data)
        {
            if (ModelState.IsValid)
            {
                // TODO

                return RedirectToAction("Index", "Home");
            }
            return View();
        }


        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Index", "Home");
        }

    }
}