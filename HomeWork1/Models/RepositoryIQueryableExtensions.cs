using System.Data.Entity.Core.Objects;
using System.Linq;


namespace HomeWork1.Models
{
	public static class RepositoryIQueryableExtensions
	{
        //如果Repository頁面有using System.Data.Entity;，會使用預設的Iclude，下述Include就會沒用
        //反之沒參考的話，就只能用下述Include
  //      public static IQueryable<T> Include<T>
		//	(this IQueryable<T> source, string path)
		//{
		//	var objectQuery = source as ObjectQuery<T>;
		//	if (objectQuery != null)
		//	{
		//		return objectQuery.Include(path);
		//	}
		//	return source;
		//}
	}
}