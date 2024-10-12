using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EH.Service.Interface.Sys.AucProductService;

namespace EH.Service.Interface.Sys
{
    public interface IAucProductService : IBaseService<Auc_Product>
    {
        public Task<string> UploadAttachments(IFormFileCollection files,string productId, string? originImgs);
        public Auc_Product DiyInsert(Auc_ProductDto dto);
        public bool Import(IFormFile file,string activityId);
        public byte[] Export(string activityId);
        public List<Auc_Product> Search(Auc_ProductSeachDto dto, out int totalCount);

        public List<Auc_MyAuctionDto> GetMyAuctions(Auc_MyAuctionRequestDto req, out int total);

    }
}
