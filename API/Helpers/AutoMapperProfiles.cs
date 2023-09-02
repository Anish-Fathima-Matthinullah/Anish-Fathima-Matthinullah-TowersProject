using API.Dtos;
using API.Modals;
using AutoMapper;

namespace API.Helpers
{
    public class AutoMapperProfiles: Profile
    {
        public AutoMapperProfiles()
        {
            CreateMap<ExcelData, ExcelDataDto>();
        }
    }
}