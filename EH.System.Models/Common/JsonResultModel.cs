using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Common
{
    public class JsonResultModel<T> 
    {
        /// <summary>
        /// 返回类
        /// </summary>
        public JsonResultModel() { }
        /// <summary>
        /// code 000成功，1XX失败，-1异常
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }
        /// <summary>
        /// 返回结果，成功不为null，失败可为null
        /// </summary>
        public T Result { get; set; }
        /// <summary>
        ///其他信息，另作他用
        /// </summary>
        public object Other { get; set; }
        /// <summary>
        /// token
        /// </summary>
        public string Token { get; set; }
    }
}
