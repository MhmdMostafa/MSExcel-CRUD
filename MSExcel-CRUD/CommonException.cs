using System;
using System.Runtime.Serialization;
namespace MSExcel_CRUD
{
    [Serializable]
    class CommonException: Exception
    {
        public CommonException()
        {
        }

        public CommonException(string message)
        : base(message)
        {
        }

        public CommonException(string message, Exception inner)
            : base(message, inner)
        {
        }
        protected CommonException(SerializationInfo info,
        StreamingContext context) : base(info, context) { }
    }
}
