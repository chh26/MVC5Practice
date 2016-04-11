using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.Entity;

namespace HomeWork1.Models
{   
	public  class 客戶聯絡人Repository : EFRepository<客戶聯絡人>, I客戶聯絡人Repository
	{
        public override IQueryable<客戶聯絡人> All()
        {
            //如果Repository頁面有using System.Data.Entity;，會使用預設的Iclude，RepositoryIQueryableExtensions的Include就會沒用
            //反之沒參考的話，就只能用RepositoryIQueryableExtensions的Include
            return base.All().Where(p => p.是否已刪除 == false).Include(客 => 客.客戶資料);
        }

        public IQueryable<客戶聯絡人> Search(string searchTxt, string jobTitle)
        {
            var data = this.All();

            if (!string.IsNullOrEmpty(searchTxt))
            {
                data = data.
                   Where(p => p.職稱.Contains(searchTxt)
                           || p.姓名.ToString().Contains(searchTxt)
                           || p.Email.ToString().Contains(searchTxt)
                           || p.手機.Contains(searchTxt)
                           || p.電話.Contains(searchTxt)
                           || p.客戶資料.客戶名稱.Contains(searchTxt)
                        ).AsQueryable();
            }

            if (!string.IsNullOrEmpty(jobTitle))
            {
                data = data.
                    Where(p => p.職稱 == jobTitle);

            }

            return data;
        }

        public 客戶聯絡人 Find(int id)
        {
            return this.All().FirstOrDefault(p => p.Id == id);
        }

        public override void Delete(客戶聯絡人 entity)
        {
            entity.是否已刪除 = true;
        }

    }

	public  interface I客戶聯絡人Repository : IRepository<客戶聯絡人>
	{

	}
}