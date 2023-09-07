using API.Dtos;

namespace API.Interfaces
{
    public interface IUpdateService
    {
        Task<bool> updateTable(ClientDataDto data);
    }
}