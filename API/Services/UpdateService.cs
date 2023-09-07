using API.Data;
using API.Interfaces;
using API.Dtos;
using AutoMapper;
using Microsoft.EntityFrameworkCore;
namespace API.Services
{
    public class UpdateService : IUpdateService
    {
        private readonly DataContext _context;
        public UpdateService(DataContext context)
        {
            _context = context;
        }

        public async Task<bool> updateTable(ClientDataDto data)
        {
            if(_context.TimeSchedule is null) return false;
            return await _context.TimeSchedule
                .Where(i => i.Id == data.Id)
                .ExecuteUpdateAsync(setters =>
                    setters.SetProperty(a => a.ActivityId, data.ActivityId)
                            .SetProperty(a => a.ActivityName, data.ActivityName)
                            .SetProperty(a => a.BlProjectStart, data.BlProjectStart == null || data.BlProjectStart.ToString() == "" ? null : Convert.ToDateTime(data.BlProjectStart))
                            .SetProperty(a => a.BlProjectFinish, data.BlProjectFinish == null || data.BlProjectFinish.ToString() == "" ? null : Convert.ToDateTime(data.BlProjectFinish))
                            .SetProperty(a => a.ActualStart, data.ActualStart == null || data.ActualStart.ToString() == "" ? null : Convert.ToDateTime(data.ActualStart))
                            .SetProperty(a => a.ActualFinish, data.ActualFinish == null || data.ActualFinish.ToString() == "" ? null : Convert.ToDateTime(data.ActualFinish))
                            .SetProperty(a => a.ActivityComplete, data.ActivityComplete == null ? null : Convert.ToDecimal(data.ActivityComplete))
                            .SetProperty(a => a.MaterialCostComplete, data.MaterialCostComplete == null ? null : Convert.ToDecimal(data.MaterialCostComplete))
                            .SetProperty(a => a.LaborCostComplete, data.LaborCostComplete == null ? null : Convert.ToDecimal(data.LaborCostComplete))
                            .SetProperty(a => a.NonLaborCostComplete, data.NonLaborCostComplete == null ? null : Convert.ToDecimal(data.NonLaborCostComplete))
                        ) > 0;
        }
    }
}