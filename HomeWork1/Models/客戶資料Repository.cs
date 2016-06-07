using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.Entity;
	
namespace HomeWork1.Models
{   
	public  class 客戶資料Repository : EFRepository<客戶資料>, I客戶資料Repository
	{
        public override IQueryable<客戶資料> All()
        {
            return base.All().Where(p => p.是否已刪除 == false).Include(path => path.客戶類別);
        }

        public 客戶資料 Find(int id)
        {
            return this.All().FirstOrDefault(p => p.Id == id);
        }

        public IQueryable<客戶資料> Search(string searchTxt, int? categoryId)
        {
            var data = this.All();

            if (!string.IsNullOrEmpty(searchTxt))
            {
                data = data.
                Where(p => p.客戶名稱.Contains(searchTxt)
                        || p.統一編號.Contains(searchTxt)
                        || p.電話.Contains(searchTxt)
                        || p.傳真.Contains(searchTxt)
                        || p.地址.Contains(searchTxt)
                        || p.Email.Contains(searchTxt)
                     );
            }

            if (categoryId.HasValue)
            {
                int intCategory = Convert.ToInt32(categoryId);
                data = data.Where(p => p.客戶分類 == intCategory);
            }

            return data;
        }

        public override void Delete(客戶資料 entity)
        {
            entity.是否已刪除 = true;

            foreach (var item in entity.客戶銀行資訊)
            {
                item.是否已刪除 = true;
            }

            foreach (var item in entity.客戶聯絡人)
            {
                item.是否已刪除 = true;
            }
        }
    }

	public  interface I客戶資料Repository : IRepository<客戶資料>
	{

	}
}