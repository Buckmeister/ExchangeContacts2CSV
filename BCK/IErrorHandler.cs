using System;

namespace BCK
{
    public interface IErrorHandler
    {
        void HandleError(Exception ex);
    }
}
