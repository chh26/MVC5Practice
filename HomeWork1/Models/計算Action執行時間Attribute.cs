using System;
using System.Web.Mvc;

namespace HomeWork1.Models
{
    public class 計算Action執行時間Attribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.ActionStart = DateTime.Now;
            base.OnActionExecuting(filterContext);
        }

        public override void OnActionExecuted(ActionExecutedContext filterContext)
        {
            filterContext.Controller.ViewBag.ActionEnd = DateTime.Now;
            TimeSpan ActionTimeSpen = filterContext.Controller.ViewBag.ActionEnd -
                filterContext.Controller.ViewBag.ActionStart;
            System.Diagnostics.Debug.WriteLine("Action執行時間：" + ActionTimeSpen.ToString());
            base.OnActionExecuted(filterContext);
        }

        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.ResultStart = DateTime.Now;
            base.OnResultExecuting(filterContext);
        }

        public override void OnResultExecuted(ResultExecutedContext filterContext)
        {
            filterContext.Controller.ViewBag.ResultEnd = DateTime.Now;
            TimeSpan ResultTimeSpen = filterContext.Controller.ViewBag.ResultEnd -
                filterContext.Controller.ViewBag.ResultStart;
            System.Diagnostics.Debug.WriteLine("ActionResult執行時間：" + ResultTimeSpen.ToString());
            base.OnResultExecuted(filterContext);
        }
    }
}