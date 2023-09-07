using API.Data;
using API.Dtos;
using API.Interfaces;
using API.Models;
using AutoMapper;

namespace API.Services
{
    public class ExportService : IExportService
    {
        private readonly DataContext _context;
        private readonly IMapper _mapper;
        public ExportService(DataContext context, IMapper mapper)
        {
            _mapper = mapper;
            _context = context;
        }

        public ExcelListDetailsDto ExportData()
        {
            if (_context.TimeSchedule is null) return new ExcelListDetailsDto();

            List<ExcelData> rowItems = _context.TimeSchedule.ToList();
            List<ExcelDataDto> rowItemsForExcel = _mapper.Map<List<ExcelDataDto>>(rowItems);

            return new ExcelListDetailsDto()
            {
                ExcelList = rowItemsForExcel,
                TableData = rowItems
            };
        }

        public List<ClientDataDto> GetData()
        {
            List<ExcelData>? tableData = ExportData().TableData;
            List<ClientDataDto> clientData = _mapper.Map<List<ClientDataDto>>(tableData);

            if (tableData == null) return new List<ClientDataDto>();
            foreach (ExcelData data in tableData)
            {
                clientData.Where(i => i.Id == data.Id).ToList().ForEach(item => {
                    string[]? style = data.Style?.Split(", ");
                    item.Color = style == null ? "#FFFFFFFF": style[0];
                    item.Font = style == null ? 11 : Convert.ToInt32(style[1]);
                    item.Bold = style == null ? false : Convert.ToBoolean(style[2]);
                });
            }

            return clientData;
        }
    }
}