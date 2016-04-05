using System;
using System.Linq;
using System.Collections.Generic;
	
namespace HomeWork1.Models
{   
	public  class View_CustomerInformStatisticsRepository : EFRepository<View_CustomerInformStatistics>, IView_CustomerInformStatisticsRepository
	{

	}

	public  interface IView_CustomerInformStatisticsRepository : IRepository<View_CustomerInformStatistics>
	{

	}
}